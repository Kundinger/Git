Attribute VB_Name = "Module22"
' Module 22  - Remote DB Manager
'
Option Explicit
Public mobjRemoteConn As ADODB.Connection
Public mobjRemoteRst As ADODB.Recordset
Public mobjRemTaskConn As ADODB.Connection
Public mobjRemTaskRst As ADODB.Recordset
Public mobjRemStatusConn As ADODB.Connection
Public mobjRemStatusRst As ADODB.Recordset
Public dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim RemoteReqTask As RemTaskControlBlock

Public Sub RemData_Clear(ByRef xREM As RemTaskControlBlock)
'
    xREM.TaskID = "na"
    xREM.VIN = "na"
    xREM.ActualShift = 0
    xREM.ActualStation = 0
    xREM.Can.WorkingVolume = 0
    xREM.Can.WorkingCapacity = 0
    xREM.InhibitChanges = False
    xREM.Rcp = EmptyRecipe       ' clear recipe values
    xREM.RequestedShift = 0
    xREM.RequestedStation = 0
    xREM.TaskStatus = "na"
    xREM.PreviousResult = "na"
    xREM.ReportComplete = False
    
End Sub

Public Function ValidRemTaskOrder(ByRef iTask As RemTaskControlBlock) As Boolean

Dim errMsg As String
Dim sMsg As String
   
    errMsg = ""
    
    ' Check new Task Order
    If (iTask.RequestedShift < 1) Then errMsg = "Invalid Requested Shift"
    If (iTask.RequestedShift > NR_SHIFT) Then errMsg = "Invalid Requested Shift"
    If (iTask.RequestedStation < 1) Then errMsg = "Invalid Requested Unit"
    If (iTask.RequestedStation > NR_STN) Then errMsg = "Invalid Requested Unit"
    If (iTask.Rcp.Number < 1) Then errMsg = "Invalid Requested Recipe"
    If (iTask.Rcp.Number > NR_RCP) Then errMsg = "Invalid Requested Recipe"
    If (iTask.Can.WorkingCapacity <= 0) Then errMsg = "Invalid Can BWC"
    If (iTask.Can.WorkingCapacity > 10000) Then errMsg = "Invalid Can BWC"
    If (iTask.Can.WorkingVolume <= 0) Then errMsg = "Invalid Can Vol"
    If (iTask.Can.WorkingVolume > 10000) Then errMsg = "Invalid Can Vol"
    If (Len(iTask.TaskID) < 3) Then errMsg = "Invalid TaskID"
    
    If (errMsg = "") Then
        ' new task order is valid
        iTask.TaskStatus = "Ready"
        ValidRemTaskOrder = True
    Else
        ' invalid new task order
        sMsg = "Task Order " & iTask.TaskID & " has " & errMsg
        Write_ELog sMsg
        iTask.TaskStatus = errMsg
        ValidRemTaskOrder = False
    End If
    
    
End Function

Public Function ValidRemRecipe(thisRcp As Integer) As Boolean
Dim errorFlag As Boolean

    errorFlag = False
    
    ' Check recipe for the station
    DispStn = CurRemoteTask.ActualStation
    DispShift = CurRemoteTask.ActualShift
    frmRecipe.Show
    frmRecipe.tmrUpdate.Enabled = True
    frmRecipe.ChgRecipeMode (STATIONMODE)
    ' Load Recipe n
    frmRecipe.LoadNewRcp thisRcp
    With frmRecipe
        .chkPrimaryScale.Value = 0
        .chkUseAuxScale = 1
        .txtPrimaryScaleNo.text = "0"
        .txtAuxScaleNo.text = Format(CurRemoteTask.ActualStation, "#0")
    End With
    If Not frmRecipe.OkToRunRecipeInStation Then
        errorFlag = True
    Else
        frmRecipe.SaveRecipe
    End If
    ValidRemRecipe = IIf(errorFlag, False, True)
End Function

Sub RemTask_Update(stn As Integer, Shift As Integer, newStatus As String, prevResult As String)
' Procedure Name:   RemTask_Update
' Created by:       MMW
' Description:      This routine updates the RemTask status
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 44

Dim Criteria, filename As String
'Dim dbDbase As Database
'Dim rsTable As Recordset
Dim msg As String
Dim srcFolder As String
Dim dstFolder As String
Dim avlFlag As Boolean
Dim iTask As RemTaskControlBlock

    If (newStatus = "Invalid") Then
    
        ' new Task Order is invalid thus no db record
        iTask = RemoteReqTask
        If (InStr(1, iTask.TaskStatus, "Invalid") = 0) Then iTask.TaskStatus = newStatus
        srcFolder = "Request"
        dstFolder = "Failed"
        avlFlag = True
        
    Else
    
        ' update Remote Task status
        iTask = StnRemoteTask(stn, Shift)
'        iTask.TaskStatus = newStatus
        ' update Remote Task status in DB
        ' update Remote Task status in DB
        ' update Remote Task status in DB
        ' Open existing Remote Task Orders Records (if any)
        Criteria = "SELECT * FROM [RemoteTasks] WHERE [RemoteTasks].[REM_TaskID] = '" & iTask.TaskID & "'"
        mobjRemTaskRst.CursorLocation = adUseClient
        mobjRemTaskRst.Open Criteria, mobjRemTaskConn, adOpenDynamic, adLockOptimistic, adCmdText
        
        With mobjRemTaskRst
            
        
            If (.BOF) Then
            
                ' no record found
                msg = "REM Task Status Update Failed for TaskID " & iTask.TaskID
                msg = msg & "  , Station " & Format(stn, "0")
                msg = msg & "  , Shift " & Format(Shift, "0")
                msg = msg & "  , - No record found)"
                Write_ELog msg
                iTask.TaskStatus = "No record found for TaskID " & iTask.TaskID
                avlFlag = IIf((newStatus = "InActive"), True, False)
                
            Else
            
                ' record(s) found
                .MoveFirst
                If (.RecordCount > 1) Then
                
                    ' more than one record found
                    msg = "REM Task Status Update Failed for TaskID " & iTask.TaskID
                    msg = msg & "  , Station " & Format(stn, "0")
                    msg = msg & "  , Shift " & Format(Shift, "0")
                    msg = msg & "  , - More than one record found)"
                    Write_ELog msg
                    iTask.TaskStatus = "More than one record found for TaskID " & iTask.TaskID
                    avlFlag = IIf((newStatus = "InActive"), True, False)
                    
                Else
                
                    ' one record found
'                    rsTable.Edit
                    Select Case newStatus
                        Case "Ready"
                            .Fields("REM_TaskStatus") = newStatus
                            .Fields("REM_PreviousResult") = prevResult
                            .Fields("REM_ActualStartDate") = 0
                            .Fields("REM_ActualDoneDate") = 0
                            StnRemoteTask(stn, Shift).PreviousResult = prevResult
                            avlFlag = False
                        Case "InProcess"
                            .Fields("REM_TaskStatus") = newStatus
                            .Fields("REM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
                            .Fields("REM_ActualStation") = stn
                            .Fields("REM_ActualShift") = Shift
                            .Fields("REM_ActualStartDate") = Now()
                            iTask.TaskStatus = newStatus
                            iTask.JobNumber = StationControl(stn, Shift).Job_Number
                            srcFolder = "OnList"
                            dstFolder = "InProcess"
                            avlFlag = True
                        Case "InActive"
                            If (.Fields("REM_TaskStatus") <> "Ready") Then
                                iTask.TaskStatus = "Cannot Inactivate - Task Status is " & .Fields("REM_TaskStatus")
                                srcFolder = "Request"
                                dstFolder = "Failed"
                                avlFlag = True
                            Else
                                If (InStr(1, UCase(iTask.TaskStatus), UCase("InActiv")) > 0) Then
                                    .Fields("REM_TaskStatus") = iTask.TaskStatus
                                Else
                                    iTask.TaskStatus = newStatus
                                    .Fields("REM_TaskStatus") = newStatus
                                End If
                                .Fields("REM_PreviousResult") = prevResult
    '                            .Fields("REM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
    '                            .Fields("REM_ActualStation") = stn
    '                            .Fields("REM_ActualShift") = Shift
    '                            .Fields("REM_ActualStartDate") = Now()
    '                            iTask.JobNumber = StationControl(stn, Shift).Job_Number
                                srcFolder = "OnList"
                                dstFolder = "Failed"
                                avlFlag = True
                            End If
                        Case "Done"
                            iTask.TaskStatus = "Done"
                            .Fields("REM_TaskStatus") = newStatus
        '                    .Fields("REM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
        '                    .Fields("REM_ActualStation") = stn
        '                    .Fields("REM_ActualShift") = Shift
                            .Fields("REM_ActualDoneDate") = Now()
                            srcFolder = "InProcess"
                            dstFolder = "Done"
                            avlFlag = True
                        Case "Failed"
                            iTask.TaskStatus = newStatus
                            .Fields("REM_TaskStatus") = newStatus
                            .Fields("REM_PreviousResult") = prevResult
        '                    .Fields("REM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
        '                    .Fields("REM_ActualStation") = stn
        '                    .Fields("REM_ActualShift") = Shift
                            .Fields("REM_ActualDoneDate") = Now()
                            srcFolder = "InProcess"
                            dstFolder = "Failed"
                            avlFlag = True
                        Case Else
                    End Select
                    .Update
                End If
                
            End If
            
            DoEvents
            
        End With
        
        StnRemoteTask(stn, Shift) = iTask
        mobjRemTaskRst.Close
        
    End If
               
    If (USINGREMAVLFILES And avlFlag) Then
        AVL_TaskFile_Move iTask, srcFolder, dstFolder
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

Sub RemTask_ClearActive()
' Procedure Name:   RemTask_ClearActive
' Created by:       MMW
' Description:      This routine clears(sets to Ready) all Active Remote Tasks
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 400

' Dim Criterion As String
' Dim dbDbase As Database
' Dim rsTable As Recordset
Dim msg As String
Dim sts As String
Dim srcFolder As String
Dim dstFolder As String
Dim iTask As RemTaskControlBlock

    sts = "InProcess"   '"Active"

    ' update Remote Task status in DB
    Criteria = "SELECT * FROM [RemoteTasks] WHERE [RemoteTasks].[REM_TaskStatus] = '" & sts & "'"
    mobjRemTaskRst.CursorLocation = adUseClient
    mobjRemTaskRst.Open Criteria, mobjRemTaskConn, adOpenDynamic, adLockOptimistic, adCmdText
    
    With mobjRemTaskRst
            
        If (.BOF) Then
            ' no tasks found; nothing to do
        Else
        
            ' task(s) found
            .MoveLast
            Do While Not .BOF
                '.Edit
                .Fields("REM_TaskStatus") = "Ready"
                .Fields("REM_PreviousResult") = "Reset from InProcess"
                .Update
                If (USINGREMAVLFILES) Then
                    srcFolder = "InProcess"
                    dstFolder = "OnList"
                    iTask.AVL_FileRoot = .Fields("AVL_FileRoot")
                    iTask.TaskID = .Fields("REM_TaskID")
                    iTask.TaskStatus = .Fields("REM_TaskStatus")
                    iTask.PreviousResult = .Fields("REM_PreviousResult")
                    iTask.VIN = .Fields("REM_VIN")
                    iTask.RequestedShift = .Fields("REM_RequestedShift")
                    iTask.RequestedStation = .Fields("REM_RequestedStation")
                    AVL_TaskFile_Move iTask, srcFolder, dstFolder
                End If
    
                .MovePrevious
            Loop
            
        End If
    
    End With
    mobjRemTaskRst.Close
          
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

Public Sub UpdateAllRemoteMasters()
'
'        Copy ALL Master
'               Canister, Recipe, PurgeProfile, & JobSequence
'               Information Records to Remote DB
'

    ' open master canister / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
    ' open remote database
    OpenConnToRemoteDb
    
    ' update remote masters
    UpdateRemoteCanisters
    UpdateRemoteRecipes
    UpdateRemotePurgeProfiles
    
    ' update remote configuration
    UpdateRemoteConfiguration
    
    ' close remote database
    CloseConnToRemoteDb
End Sub
    
Public Sub OpenConnToRemoteDb()
'
'
    ' Connect to Remote Database
    Set mobjRemoteConn = New ADODB.Connection
    Set mobjRemoteRst = New ADODB.Recordset
    mobjRemoteConn.ConnectionString = "DSN=cpsRemote;Uid=;Pwd=;"
    mobjRemoteConn.Open
    
End Sub

Public Sub CloseConnToRemoteDb()
'
'
    ' Close Connection to Remote Database
    mobjRemoteConn.Close
    
End Sub

Public Sub OpenConnToRemStatusDb()
'
'
    ' Connect to Remote Status Database
    Set mobjRemStatusConn = New ADODB.Connection
    Set mobjRemStatusRst = New ADODB.Recordset
    mobjRemStatusConn.ConnectionString = "DSN=cpsRemote;Uid=;Pwd=;"
    mobjRemStatusConn.Open
    
End Sub

Public Sub CloseConnToRemStatusDb()
'
'
    ' Close Connection to Remote Status Database
    mobjRemStatusConn.Close
    
End Sub

Public Sub OpenConnToRemTaskDb()
'
'
    ' Connect to Remote Task Database
    Set mobjRemTaskConn = New ADODB.Connection
    Set mobjRemTaskRst = New ADODB.Recordset
    mobjRemTaskConn.ConnectionString = "DSN=cpsRemote;Uid=;Pwd=;"
    mobjRemTaskConn.Open
    
End Sub

Public Sub CloseConnToRemTaskDb()
'
'
    ' Close Connection to Remote Task Database
    mobjRemTaskConn.Close
    
End Sub

Public Sub UpdateRemoteCanisters()
'
'        Copy Master Canister Information Records to Remote DB
'
Dim iCan As Integer
        
    ' CANISTERS
    ' CANISTERS
    ' CANISTERS
    For iCan = 1 To MAX_CANRCP
    
        ' Read Master Canister Information Record
        Criteria = "SELECT * FROM [MasterCanister] WHERE [Number] = " & iCan & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If Not rsRecord.BOF Then
           
            ' Open existing Remote Master Canister Information Record (if any)
            Criteria = "SELECT * FROM [MasterCanister] WHERE [MasterCanister].[Number] = " & iCan & " "
            mobjRemoteRst.CursorLocation = adUseClient
            mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
            
            With mobjRemoteRst
            
                If .BOF Then
                    .AddNew
                    .Fields("Number").Value = iCan
                Else
                    .MoveLast
                    .MoveFirst
                End If
                   
                Select Case .RecordCount
                    Case 1
                        ' Update Remote Master Canister Information Record
                        .Fields("Description").Value = rsRecord.Fields("Description").Value
                        .Fields("WorkingCapacity").Value = rsRecord.Fields("WorkingCapacity").Value
                        .Fields("WCVolume").Value = rsRecord.Fields("WCVolume").Value
                        .Update
                        DoEvents
                    Case Is > 1
                        ' Error - Multiple Records Returned
                        Write_ELog "RemoteCan Update Failure - Multiple Records Returned for Can# " & Format(iCan, "#,##0")
                End Select
                
            End With
            mobjRemoteRst.Close
        
        End If

        rsRecord.Close
        
    Next iCan
        
    If Not VarInitDone Then frmAbout.UpdateMsg "Updated Remote Master Canisters" & vbCrLf
    Write_ELog "Updated Remote Master Canisters"
    
    
End Sub

Public Sub UpdateRemoteRecipes()
'
'        Copy Master Recipe Information Records to Remote DB
'
Dim iRcp As Integer
    
    ' RECIPES
    ' RECIPES
    ' RECIPES
    For iRcp = 1 To MAX_RCP
    
        ' Read Master Recipe Information Record
        Criteria = "SELECT * FROM [MasterRecipe] WHERE [Number] = " & iRcp & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If Not rsRecord.BOF Then
                       
            ' Open existing Remote Master Recipe Information Record (if any)
            Criteria = "SELECT * FROM [MasterRecipe] WHERE [MasterRecipe].[Number] = " & iRcp & " "
            mobjRemoteRst.CursorLocation = adUseClient
            mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
            
            With mobjRemoteRst
            
                If .BOF Then
                    .AddNew
                    .Fields("Number").Value = iRcp
                Else
                    .MoveLast
                    .MoveFirst
                End If
                   
                Select Case .RecordCount
                    Case 1
                        ' Update Remote Master Recipe Information Record
                        .Fields("Name").Value = rsRecord.Fields("Name").Value
                        
                        .Fields("CycleType").Value = rsRecord.Fields("CycleType").Value
                        .Fields("CycleTypeDesc").Value = CycleTypeDesc(rsRecord.Fields("CycleType").Value)
                                
                        .Fields("Load_Method").Value = rsRecord.Fields("Load_Method").Value
                        .Fields("Load_MethodDesc").Value = LoadMethodDesc(rsRecord.Fields("Load_Method").Value)
                        .Fields("NitrogenFlow").Value = rsRecord.Fields("NitrogenFlow").Value
                        .Fields("Load_Rate").Value = rsRecord.Fields("Load_Rate").Value
                        .Fields("Mix_Percent").Value = rsRecord.Fields("Mix_Percent").Value
                        .Fields("WC_Mult").Value = rsRecord.Fields("WC_Mult").Value
                        .Fields("EPAFill").Value = rsRecord.Fields("EPAFill").Value
                        .Fields("Load_Wt").Value = rsRecord.Fields("Load_Wt").Value
                        .Fields("LoadBreakthrough").Value = rsRecord.Fields("LoadBreakthrough").Value
                        .Fields("FIDmg").Value = rsRecord.Fields("FIDmg").Value
                        .Fields("Load_Time").Value = rsRecord.Fields("Load_Time").Value
                        .Fields("Purge_Method").Value = rsRecord.Fields("Purge_Method").Value
                        .Fields("Purge_MethodDesc").Value = PurgeMethodDesc(rsRecord.Fields("Purge_Method").Value)
                        .Fields("Purge_AuxTime").Value = rsRecord.Fields("Purge_AuxTime").Value
                        .Fields("Purge_Time").Value = rsRecord.Fields("Purge_Time").Value
                        .Fields("Purge_Flow").Value = rsRecord.Fields("Purge_Flow").Value
                        .Fields("Purge_Can_Vol").Value = rsRecord.Fields("Purge_Can_Vol").Value
                        .Fields("Purge_ProfileNumber").Value = rsRecord.Fields("Purge_ProfileNumber").Value
                        .Fields("Purge_TargetMode").Value = rsRecord.Fields("Purge_TargetMode").Value
                        .Fields("Purge_TargetModeDesc").Value = PurgeTargetDesc(rsRecord.Fields("Purge_TargetMode").Value)
                        .Fields("Purge_TargetWC").Value = rsRecord.Fields("Purge_TargetWC").Value
                        .Fields("Purge_TargetWeight").Value = rsRecord.Fields("Purge_TargetWeight").Value
                        .Fields("Purge_MaxVolumes").Value = rsRecord.Fields("Purge_MaxVolumes").Value
                        .Fields("Purge_TargetPurge").Value = rsRecord.Fields("Purge_TargetPurge").Value
                        .Fields("Purge_TargetPause").Value = rsRecord.Fields("Purge_TargetPause").Value
                        
                        .Fields("PurgeAuxCan").Value = rsRecord.Fields("PurgeAuxCan").Value
                        .Fields("PurgeCansInSeries").Value = rsRecord.Fields("PurgeCansInSeries").Value
                        .Fields("UseAuxScale").Value = rsRecord.Fields("UseAuxScale").Value
                        .Fields("AuxScaleNo").Value = rsRecord.Fields("AuxScaleNo").Value
                        .Fields("PauseLeakTime").Value = rsRecord.Fields("PauseLeakTime").Value
                        .Fields("PauseLoadTime").Value = rsRecord.Fields("PauseLoadTime").Value
                        .Fields("PausePurgeTime").Value = rsRecord.Fields("PausePurgeTime").Value
                        .Fields("UsePriScale").Value = rsRecord.Fields("UsePriScale").Value
                        .Fields("PriScaleNo").Value = rsRecord.Fields("PriScaleNo").Value
                        .Fields("PauseAfterLeak").Value = rsRecord.Fields("PauseAfterLeak").Value
                        .Fields("PauseAfterLoad").Value = rsRecord.Fields("PauseAfterLoad").Value
                        .Fields("PauseAfterPurge").Value = rsRecord.Fields("PauseAfterPurge").Value
'                        .Fields("TargetConcentration").Value = rsRecord.Fields("TargetConcentration").Value
'                        .Fields("DwellTime").Value = rsRecord.Fields("DwellTime").Value
                        .Fields("LeakCheck").Value = rsRecord.Fields("LeakCheck").Value
                        .Fields("LeakPrimary").Value = rsRecord.Fields("LeakPrimary").Value
                        .Fields("LeakAux").Value = rsRecord.Fields("LeakAux").Value
'                        .Fields("UseAnalyzer").Value = rsRecord.Fields("UseAnalyzer").Value
                        .Fields("MaxLoadTime").Value = rsRecord.Fields("MaxLoadTime").Value
                        .Fields("UseHiRangeMFC").Value = rsRecord.Fields("UseHiRangeMFC").Value
                        .Fields("UseLoadRatePID").Value = rsRecord.Fields("UseLoadRatePID").Value
                        
                        .Fields("IDLoad").Value = rsRecord.Fields("IDLoad").Value
                        .Fields("LoadL").Value = rsRecord.Fields("LoadL").Value
                        .Fields("LoadV").Value = rsRecord.Fields("LoadV").Value
                        .Fields("IDPurge").Value = rsRecord.Fields("IDPurge").Value
                        .Fields("PurgeL").Value = rsRecord.Fields("PurgeL").Value
                        .Fields("PurgeV").Value = rsRecord.Fields("PurgeV").Value
                        .Fields("IDVent").Value = rsRecord.Fields("IDVent").Value
                        .Fields("VentL").Value = rsRecord.Fields("VentL").Value
                        .Fields("VentV").Value = rsRecord.Fields("VentV").Value
                        
                        .Fields("LiveFuel").Value = rsRecord.Fields("LiveFuel").Value
                        .Fields("LiveFuelChgAuto").Value = rsRecord.Fields("LiveFuelChgAuto").Value
                        .Fields("LiveFuelChgFreq").Value = rsRecord.Fields("LiveFuelChgFreq").Value
                        .Fields("ADF_Heater").Value = rsRecord.Fields("ADF_Heater").Value
                        .Fields("ADF_HeaterSP").Value = rsRecord.Fields("ADF_HeaterSP").Value
                        
                        ' start method
                        .Fields("StartMethod").Value = rsRecord.Fields("StartMethod").Value
                        .Fields("StartMethodDesc").Value = StartMethodDesc(rsRecord.Fields("StartMethod").Value)
                        .Fields("StartDelay").Value = rsRecord.Fields("StartDelay").Value
                        .Fields("StartDate").Value = rsRecord.Fields("StartDate").Value
                            
                        ' end method
                        .Fields("EndMethod").Value = rsRecord.Fields("EndMethod").Value
                        .Fields("EndMethodDesc").Value = EndMethodDesc(rsRecord.Fields("EndMethod").Value)
                        .Fields("EndMaximumCycles").Value = rsRecord.Fields("EndMaximumCycles").Value
                        .Fields("EndMinimumCycles").Value = rsRecord.Fields("EndMinimumCycles").Value
                        .Fields("EndConsecutiveCycles").Value = rsRecord.Fields("EndConsecutiveCycles").Value
                        .Fields("EndWeightTolerance").Value = rsRecord.Fields("EndWeightTolerance").Value
                        .Fields("UpdateCanWc").Value = rsRecord.Fields("UpdateCanWc").Value
                        .Fields("Cycles").Value = rsRecord.Fields("Cycles").Value
                            
                        ' aux outputs
                        .Fields("AuxOutputs").Value = rsRecord.Fields("AuxOutputs").Value
                        .Fields("AuxOutput1_Load").Value = rsRecord.Fields("AuxOutput1_Load").Value
                        .Fields("AuxOutput1_Purge").Value = rsRecord.Fields("AuxOutput1_Purge").Value
                        .Fields("AuxOutput2_Load").Value = rsRecord.Fields("AuxOutput2_Load").Value
                        .Fields("AuxOutput2_Purge").Value = rsRecord.Fields("AuxOutput2_Purge").Value
                        .Fields("AuxOutput3_Load").Value = rsRecord.Fields("AuxOutput3_Load").Value
                        .Fields("AuxOutput3_Purge").Value = rsRecord.Fields("AuxOutput3_Purge").Value
                        .Fields("AuxOutput4_Load").Value = rsRecord.Fields("AuxOutput4_Load").Value
                        .Fields("AuxOutput4_Purge").Value = rsRecord.Fields("AuxOutput4_Purge").Value
                        .Update
                        DoEvents
                    Case Is > 1
                        Write_ELog "RemoteRcp Update Failure - Multiple Records Returned for Rcp# " & Format(iRcp, "#,##0")
                End Select
                               
            End With
            mobjRemoteRst.Close
        
        End If

        rsRecord.Close
        
    Next iRcp
        
    If Not VarInitDone Then frmAbout.UpdateMsg "Updated Remote Master Recipes" & vbCrLf
    Write_ELog "Updated Remote Master Recipes"
    
End Sub

Public Sub UpdateRemotePurgeProfiles()
'
'        Copy Master PurgeProfile Information Records to Remote DB
'
Dim iPrf As Integer
    
    ' PURGEPROFILES
    ' PURGEPROFILES
    ' PURGEPROFILES
    For iPrf = 1 To MAX_PROFILES
    
        ' Read Master Profiles Information Record
        Criteria = "SELECT * FROM [MasterProfiles] WHERE [Number] = " & iPrf & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        If Not rsRecord.BOF Then
           
            ' Open existing Remote Master Recipe Information Record (if any)
            Criteria = "SELECT * FROM [MasterProfiles] WHERE [MasterProfiles].[Number] = " & iPrf & " "
            mobjRemoteRst.CursorLocation = adUseClient
            mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
            
            With mobjRemoteRst
            
                If .BOF Then
                    .AddNew
                    .Fields("Number").Value = iPrf
                Else
                    .MoveLast
                    .MoveFirst
                End If
                   
                Select Case .RecordCount
                    Case 1
                        ' Update Remote Master Recipe Information Record
                        .Fields("Description").Value = rsRecord.Fields("Description").Value
                        .Fields("Steps").Value = rsRecord.Fields("Steps").Value
                        .Fields("TotalDuration").Value = rsRecord.Fields("TotalDuration").Value
                        .Fields("ProjectedLiters").Value = rsRecord.Fields("ProjectedLiters").Value
                        .Fields("ProjectedVolumes").Value = rsRecord.Fields("ProjectedVolumes").Value
                        .Update
                        DoEvents
                    Case Is > 1
                        Write_ELog "RemotePrf Update Failure - Multiple Records Returned for Profile # " & Format(iPrf, "#,##0")
                End Select
                               
'                .Close
                
            End With
            mobjRemoteRst.Close
        
        End If

        rsRecord.Close
        
        ' Save Master PurgeProfile Steps
        Criteria = "SELECT * FROM [MasterProfileSteps] WHERE [ProfileNumber] = " & iPrf & " ORDER BY [StepNumber] ASC"
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        ' Open existing Remote Master Profile Steps Information Record (if any)
        Criteria = "SELECT * FROM [MasterProfileSteps] WHERE [MasterProfileSteps].[ProfileNumber] = " & iPrf & " ORDER BY [StepNumber] ASC"
        mobjRemoteRst.CursorLocation = adUseClient
        mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
            
        If Not rsRecord.BOF Then
           
            With mobjRemoteRst
            
                ' first remove existing steps
                If .BOF Then
                    ' nothing to do; no steps for this profile exist in db
                Else
                    .MoveLast
                    If (Not .BOF) Then
                        While Not .BOF
                            .Delete
                            .MovePrevious
                        Wend
                    End If
                End If
                
                ' now add the "new" steps
                rsRecord.MoveLast
                rsRecord.MoveFirst
                Do While rsRecord.EOF
                
                    .AddNew
                    .Fields("ProfileNumber").Value = iPrf
                    .Fields("StepNumber").Value = rsRecord.Fields("StepNumber").Value
                    .Fields("InitialSP").Value = rsRecord.Fields("InitialSP").Value
                    .Fields("Duration").Value = rsRecord.Fields("Duration").Value
                    .Fields("StepType").Value = rsRecord.Fields("StepType").Value
                    .Fields("StepTypeDesc").Value = rsRecord.Fields("StepTypeDesc").Value
                    .Update
                
                    If Not rsRecord.EOF Then rsRecord.MoveNext
                    
                Loop
                    
            End With
            mobjRemoteRst.Close
            
        Else
        
            With mobjRemoteRst
            
                ' remove any existing steps
                If .BOF Then
                    ' nothing to do; no steps for this profile exist in db
                Else
                    .MoveLast
                    If (Not .BOF) Then
                        While Not .BOF
                            .Delete
                            .MovePrevious
                        Wend
                    End If
                End If
                
            End With
            mobjRemoteRst.Close
            
        End If

        rsRecord.Close
        
    Next iPrf
    
    If Not VarInitDone Then frmAbout.UpdateMsg "Updated Remote Master PurgeProfiles" & vbCrLf
    Write_ELog "Updated Remote Master PurgeProfiles"
    
End Sub

Public Sub UpdateRemoteConfiguration()
'
'        Copy Configuration Information Records to Remote DB
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 90
        
    If Not VarInitDone Then frmAbout.UpdateMsg "Begin Update of Remote Configuration" & vbCrLf
    ' CONFIGURATION
    ' CONFIGURATION
    ' CONFIGURATION
    ' Open existing Remote Configuration Information Record (if any)
    Criteria = "SELECT * FROM [Configuration] "
    mobjRemoteRst.CursorLocation = adUseClient
    mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
    
    With mobjRemoteRst
    
        If .BOF Then
            .AddNew
        Else
            .MoveLast
            .MoveFirst
        End If
           
ChgErrModule 22, 91
        Select Case .RecordCount
            Case 1
ChgErrModule 22, 92
                ' Update Remote Configuration Information Record
                .Fields("UpdateDts").Value = Now()
                .Fields("Heading").Value = SysConfig.Heading
                .Fields("Heading2").Value = SysConfig.Heading2
                .Fields("Next_File").Value = SysConfig.Next_File
                .Fields("AutoLogon").Value = SysConfig.AutoLogon
                .Fields("AutoLogonUser").Value = SysConfig.AutoLogonUser
                .Fields("DbFileBackup_Active").Value = SysConfig.DbFileBackup_Active
                .Fields("DbFileBackup_Path").Value = SysConfig.DbFileBackup_Path
                .Fields("ReportBackup_Active").Value = SysConfig.ReportBackup_Active
                .Fields("ReportBackup_Path").Value = SysConfig.ReportBackup_Path
                .Fields("EventRecs").Value = SysConfig.EventRecs
                .Fields("JobRecs").Value = SysConfig.JobRecs
                .Fields("LCMinDelay").Value = SysConfig.LCMinDelay
                .Fields("LCSetPoint").Value = SysConfig.LCSetPoint
                .Fields("LCTime").Value = SysConfig.LCTime
                .Fields("PressureDecay").Value = SysConfig.PressureDecay
                .Fields("LeakCheckFailResponse").Value = SysConfig.LeakCheckFailResponse
                .Fields("NitrogenPurgeTime").Value = SysConfig.NitrogenPurgeTime
                .Fields("CanVent_Delay_Max").Value = SysConfig.CanVent_Delay_Max
                .Fields("OOTtimeDelay").Value = SysConfig.OOTtimeDelay
                .Fields("PosPressPurge").Value = SysConfig.PosPressPurge
                .Fields("DoorOpenDelay").Value = SysConfig.DoorOpenDelay
                .Fields("UPSOpenDelay").Value = SysConfig.UPSOpenDelay
                .Fields("LoadPressure").Value = SysConfig.LoadPressure
                .Fields("ButaneMassLimit").Value = SysConfig.ButaneMassLimit
                .Fields("LoadTimeLimit").Value = SysConfig.LoadTimeLimit
                .Fields("WaterBathControl").Value = SysConfig.WaterBathControl
                .Fields("LeakCheck_Interval").Value = SysConfig.LeakCheck_Interval
                .Fields("LeakTotal_Interval").Value = SysConfig.LeakTotal_Interval
                .Fields("Load_Interval").Value = SysConfig.Load_Interval
                .Fields("Purge_Interval").Value = SysConfig.Purge_Interval
                .Fields("LoadTotal_Interval").Value = SysConfig.LoadTotal_Interval
                .Fields("PurgeTotal_Interval").Value = SysConfig.PurgeTotal_Interval
                .Fields("Tol_Nit_Flow").Value = SysConfig.Tol_Nit_Flow
                .Fields("Tol_Btn_Flow").Value = SysConfig.Tol_Btn_Flow
                .Fields("Tol_ORVRNit_Flow").Value = SysConfig.Tol_ORVRNit_Flow
                .Fields("Tol_ORVRBtn_Flow").Value = SysConfig.Tol_ORVRBtn_Flow
                .Fields("Tol_Pur_Flow").Value = SysConfig.Tol_Pur_Flow
                .Fields("Tol_Lfv_Flow").Value = SysConfig.Tol_Lfv_Flow
                .Fields("Tol_Mix_Ratio").Value = SysConfig.Tol_Mix_Ratio
                .Fields("Tol_Temp").Value = SysConfig.Tol_Temp
                .Fields("Tol_Moisture").Value = SysConfig.Tol_Moisture
                .Fields("Tol_FuelTemp").Value = SysConfig.Tol_FuelTemp
                .Fields("Tol_Purge_Total").Value = SysConfig.Tol_Purge_Total
                .Fields("Tol_Load_Total").Value = SysConfig.Tol_Load_Total
                .Fields("Tol_PurgeOven").Value = SysConfig.Tol_PurgeOvenTemp
                .Fields("Tol_WaterBath").Value = SysConfig.Tol_WaterBathTemp
                .Fields("LoLim_Load_Flow").Value = SysConfig.LoLim_Load_Flow
                .Fields("LoLim_Purge_Flow").Value = SysConfig.LoLim_Purge_Flow
                .Fields("PurgeDP_HiLimit").Value = SysConfig.PurgeDP_HiLimit
                .Fields("Temp_Target").Value = SysConfig.Temp_Target
                .Fields("Moisture_Target").Value = SysConfig.Moisture_Target
                .Fields("LoadSettleTime").Value = SysConfig.LoadSettleTime
                .Fields("PurgeSettleTime").Value = SysConfig.PurgeSettleTime
                .Fields("ReportFileName1stPart").Value = SysConfig.ReportFileName1stPart
                .Fields("ReportFileName2ndPart").Value = SysConfig.ReportFileName2ndPart
                .Fields("ReportFileName3rdPart").Value = SysConfig.ReportFileName3rdPart
            
                .Fields("CsvEotReporting").Value = SysConfig.RptConfig.CsvEotReporting
                .Fields("CsvEotSummary").Value = SysConfig.RptConfig.CsvEotSummary
                .Fields("CsvEotDetail").Value = SysConfig.RptConfig.CsvEotDetail
                .Fields("CsvGenReporting").Value = SysConfig.RptConfig.CsvGenReporting
                .Fields("CsvGenSummary").Value = SysConfig.RptConfig.CsvGenSummary
                .Fields("CsvGenDetail").Value = SysConfig.RptConfig.CsvGenDetail
                .Fields("TextEotReporting").Value = SysConfig.RptConfig.TextEotReporting
                .Fields("TextEotSummary").Value = SysConfig.RptConfig.TextEotSummary
                .Fields("TextEotSummary_AutoPrint").Value = SysConfig.RptConfig.TextEotSummary_AutoPrint
                .Fields("TextEotDetail").Value = SysConfig.RptConfig.TextEotDetail
                .Fields("TextGenReporting").Value = SysConfig.RptConfig.TextGenReporting
                .Fields("TextGenSummary").Value = SysConfig.RptConfig.TextGenSummary
                .Fields("TextGenDetail").Value = SysConfig.RptConfig.TextGenDetail
                .Fields("XlsEotReporting").Value = SysConfig.RptConfig.XlsEotReporting
                .Fields("XlsEotSummary").Value = SysConfig.RptConfig.XlsEotSummary
                .Fields("XlsEotDetail").Value = SysConfig.RptConfig.XlsEotDetail
                .Fields("XlsGenReporting").Value = SysConfig.RptConfig.XlsGenReporting
                .Fields("XlsGenSummary").Value = SysConfig.RptConfig.XlsGenSummary
                .Fields("XlsGenDetail").Value = SysConfig.RptConfig.XlsGenDetail
            
ChgErrModule 22, 94
                .Fields("BtnFlowResponse").Value = SysConfig.BtnFlowResp
                .Fields("NitFlowResponse").Value = SysConfig.NitFlowResp
                .Fields("FuelLevelResponse").Value = SysConfig.FuelLevelResp
                .Fields("FuelTempResponse").Value = SysConfig.FuelTempResp
                .Fields("PurFlowResponse").Value = SysConfig.PurFlowResp
                .Fields("AirMoistResponse").Value = SysConfig.AirMoistResp
                .Fields("AirTempResponse").Value = SysConfig.AirTempResp
                .Fields("CanVentResponse").Value = SysConfig.CanVentResp
                .Fields("LoadRateResponse").Value = SysConfig.LoadRateResp
ChgErrModule 22, 95
                .Fields("PurgeDpResponse").Value = SysConfig.PurgeDpResp
                .Fields("StorageLevelResponse").Value = SysConfig.StorageLevelResp
                .Fields("PurgeOvenResponse").Value = SysConfig.PurgeOvenResp
                .Fields("WaterBathResponse").Value = SysConfig.WaterBathResp
ChgErrModule 22, 98
                .Update
                DoEvents
            Case Is > 1
                ' Error - Multiple Records Returned
                Write_ELog "RemoteCfg Update Failure - Multiple Records Returned for Configuration"
        End Select
        
ChgErrModule 22, 99
    End With
    mobjRemoteRst.Close
    
         
    If Not VarInitDone Then frmAbout.UpdateMsg "Updated Remote Configuration" & vbCrLf
    Write_ELog "Updated Remote Master Configuration"
    
    
    ' SYSDEF
    ' SYSDEF
    ' SYSDEF
    ' Open existing Remote Sysdef Information Record (if any)
    Criteria = "SELECT * FROM [SysDefMain] "
    mobjRemoteRst.CursorLocation = adUseClient
    mobjRemoteRst.Open Criteria, mobjRemoteConn, adOpenDynamic, adLockOptimistic, adCmdText
    
    With mobjRemoteRst
    
        If .BOF Then
            .AddNew
        Else
            .MoveLast
            .MoveFirst
        End If
           
        Select Case .RecordCount
            Case 1
                ' Update Remote Sysdef Information Record
                .Fields("UsingC").Value = USINGC
                .Fields("UsingF").Value = USINGF
                .Fields("UsingMoist_RH").Value = USINGMoist_RH
                .Fields("UsingMoist_Grains").Value = USINGMoist_Grains
                .Fields("UsingLV_English").Value = USINGLVol_Engl
                .Fields("UsingLV_SI").Value = USINGLVol_SI
                .Update
                DoEvents
            Case Is > 1
                ' Error - Multiple Records Returned
                Write_ELog "RemoteSysdef Update Failure - Multiple Records Returned for Sysdef"
        End Select
    
    End With
    mobjRemoteRst.Close
    
         
    If Not VarInitDone Then frmAbout.UpdateMsg "Updated Remote System Definition" & vbCrLf
    Write_ELog "Updated Remote System Definition"
    
    
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

Public Sub AVL_TaskFiles_Check()
'
'        Validate and Error Check any new AVL Task Request Files
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 31590
Dim srcFilePath As String
Dim dstFilePath As String
Dim iSrcFileNumber As Integer
Dim iDstFileNumber As Integer
Dim a As Integer
Dim b As Integer
Dim charPos As Long
Dim tempstr As String
Dim tempstr2 As String
Dim tempdate As Date
Dim bDone As Boolean
Dim filList() As fList
Dim sDir As String
Dim sFile As String
Dim sPath As String
Dim BFile As String
Dim CFile As String
Dim dDir As String
Dim dFile As String
Dim dPath As String
Dim strExists As String
Dim linestr As String
Dim descstr As String
Dim valuestr As String
Dim chrnumEquals As Integer
Dim sCanVol As Single
Dim sCanWC As Single
Dim sCanDesc As String
Dim sRcpText As String
Dim iRcpNum As Integer
Dim sRcpName As String
Dim sTaskID As String
Dim sVIN As String
Dim sReqUnitText As String
Dim iReqStn As Integer
Dim iReqShift As Integer
Dim bInhibitChgs As Boolean
Dim fs, f As Object
Set fs = CreateObject("Scripting.FileSystemObject")
        
'  FILEPATH_avltasks

ReDim filList(0)
    sDir = FILEPATH_avltasks & "Request\" & "*.*"
    sFile = Dir(sDir)
    sPath = FILEPATH_avltasks & "Request\"
    
'Debug.Print "Looking for new Task Orders"
    frmMainMenu.MousePointer = vbHourglass
    frmDelayBox.MousePointer = vbHourglass

    ' Create array of files in the Requests directory
    Do While sFile <> ""
        ' make sure it is a data file
        If InStr(UCase$(sFile), UCase$(Trim$("TXT"))) <> 0 Then
            ' Expand array to hold this file
'            sFile = sPath + sFile
Debug.Print "Found File = " & sPath & sFile
            filList(UBound(filList)).fName = sFile
            filList(UBound(filList)).fPath = sPath
            filList(UBound(filList)).fDate = FileDateTime(sPath & sFile)
            ReDim Preserve filList(UBound(filList) + 1)
        End If
        sFile = Dir
    Loop
    
'For a = 0 To UBound(filList) - 1
'   Debug.Print "1File = " & filList(a).fName & "  " & filList(a).fDate
'Next a
    
    ' Bubble sort the files
    For a = 0 To UBound(filList) - 2
        For b = a + 1 To UBound(filList) - 1
            If filList(a).fDate > filList(b).fDate Then
                tempstr = filList(a).fName
                tempstr2 = filList(a).fPath
                tempdate = filList(a).fDate
                filList(a).fName = filList(b).fName
                filList(a).fPath = filList(b).fPath
                filList(a).fDate = filList(b).fDate
                filList(b).fName = tempstr
                filList(b).fPath = tempstr2
                filList(b).fDate = tempdate
            End If
        Next b
    Next a
    
'For a = 0 To UBound(filList) - 1
'   Debug.Print "2File = " & filList(a).fPath & filList(a).fName
'Next a
    
    For a = 0 To UBound(filList) - 1
Debug.Print "Check File = " & filList(a).fPath & filList(a).fName
        RemData_Clear RemoteReqTask
        RemoteReqTask.AVL_FileRoot = filList(a).fName
        iSrcFileNumber = FreeFile
        Open filList(a).fPath & filList(a).fName For Input As #iSrcFileNumber
        bDone = False
        Do While ((Not bDone) And (Not EOF(iSrcFileNumber)))
ChgErrModule 22, 31591
            Input #iSrcFileNumber, linestr
ChgErrModule 22, 31592
'Debug.Print "Line - " & linestr
            chrnumEquals = InStr(1, linestr, "=")
            descstr = Mid(linestr, 1, (chrnumEquals - 1))
            valuestr = Mid(linestr, (chrnumEquals + 1), (Len(linestr) - chrnumEquals))
'Debug.Print "Dsc - " & descstr & " <><><> " & "Val - " & valuestr
            Select Case Trim(descstr)
                Case "TaskId"
                    If (Len(valuestr) > 3) Then RemoteReqTask.TaskID = valuestr
                Case "TaskCaption"
                    RemoteReqTask.TaskCaption = valuestr
                Case "VINNumber"
                    RemoteReqTask.VIN = valuestr
                Case "CanisterDescription"
                    RemoteReqTask.Can.Description = valuestr
                Case "CanisterVolume"
                    RemoteReqTask.Can.WorkingVolume = ValueFromText(valuestr)
                Case "WorkingCapacity"
                    RemoteReqTask.Can.WorkingCapacity = ValueFromText(valuestr)
                Case "Recipe"
                    If (Len(valuestr) > 1) Then
                        charPos = InStr(1, valuestr, "-")
                        If (charPos > 0) Then
                            RemoteReqTask.Rcp.Number = CInt(Trim(Mid(valuestr, 1, (charPos - 1))))
                            RemoteReqTask.Rcp.Name = Trim(Mid(valuestr, (charPos + 1), (Len(valuestr) - charPos)))
                        Else
                            RemoteReqTask.Rcp.Number = 0
                            RemoteReqTask.Rcp.Name = "none"
                        End If
                    Else
                        RemoteReqTask.Rcp.Number = 0
                        RemoteReqTask.Rcp.Name = "none"
                    End If
                Case "RequestedBaseUnitAndShift"
                    If (Len(valuestr) > 1) Then
                        RemoteReqTask.RequestedStation = LocalStnNumfromSysID(Mid(valuestr, 1, (Len(valuestr) - 1)))
                        RemoteReqTask.RequestedShift = CInt(ValueFromText(Mid(valuestr, Len(valuestr), 1)))
                    Else
                        RemoteReqTask.RequestedStation = 0
                        RemoteReqTask.RequestedShift = 0
                    End If
                Case "ActualBaseUnitAndShift"
                    RemoteReqTask.ActualStation = 0
                    RemoteReqTask.ActualShift = 0
                Case "JobNumber"
                    RemoteReqTask.JobNumber = valuestr
                Case "Status"
'                    newTask.TaskStatus = "Ready"
                    valuestr = Trim(valuestr)
'                    If (Len(valuestr) > 3) Then
'                        Select Case valuestr
'                            Case "ReqRun"
                                RemoteReqTask.TaskStatus = "Ready"
'                            Case "ReqIna"
'                                RemoteReqTask.TaskStatus = "InActive"
'                            Case Else
'                                RemoteReqTask.TaskStatus = "Invalid Status"
'                        End Select
'                    Else
'                        RemoteReqTask.TaskStatus = "Invalid Status"
'                    End If
                    bDone = True
            End Select
        Loop
        If (EOF(iSrcFileNumber) And (Not bDone)) Then RemoteReqTask.TaskStatus = "Invalid - Missing Status"
        Close #iSrcFileNumber
        ' chk new Task then add
        Select Case RemoteReqTask.TaskStatus
            Case "Ready"
                ' new Task Order is "Ready"
                If (ValidRemTaskOrder(RemoteReqTask)) Then
                    AddNewTask RemoteReqTask
                Else
                   ' new Task Order is "Invalid"
                   RemTask_Update 0, 0, "Invalid", "none"
                End If
            Case "InActive"
                ' existing Task Order is "InActive"
                StnRemoteTask(0, 0) = RemoteReqTask
                StnRemoteTask(0, 0).TaskStatus = "Inactivated by Host"
                RemTask_Update 0, 0, "InActive", "Inactivated by Host"
            Case Else
                ' new Task Order is "Invalid"
                RemTask_Update 0, 0, "Invalid", "none"
        End Select
    
    Next a
    
    frmMainMenu.MousePointer = vbDefault
    frmDelayBox.MousePointer = vbDefault
    
    
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

Public Function LocalStnNumfromSysID(ByVal sID As String) As Integer
Dim iStn As Integer
Dim Idx As Integer
    
    Idx = 0
    If (Len(sID) > 1) Then
        For iStn = 1 To NR_STN
            If (sID = STN_INFO(iStn).SysID) Then Idx = iStn
        Next iStn
    End If
    LocalStnNumfromSysID = Idx
    
End Function

Public Sub AddNewTask(ByRef iTask As RemTaskControlBlock)
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 31591
Dim errFlag As Boolean

            ' Open existing Remote Task Orders Records (if any)
            errFlag = False
            Criteria = "SELECT * FROM [RemoteTasks] WHERE [RemoteTasks].[REM_TaskID] = '" & iTask.TaskID & "'"
            mobjRemTaskRst.CursorLocation = adUseClient
            mobjRemTaskRst.Open Criteria, mobjRemTaskConn, adOpenDynamic, adLockOptimistic, adCmdText
            
            With mobjRemTaskRst
            
                If .BOF Then
                    .AddNew
                    .Fields("REM_TaskID").Value = iTask.TaskID
                    .Fields("REM_VIN").Value = iTask.VIN
                    .Fields("REM_RequestedStation").Value = iTask.RequestedStation
                    .Fields("REM_RequestedShift").Value = iTask.RequestedShift
                    .Fields("REM_ActualStation").Value = 0
                    .Fields("REM_ActualShift").Value = 0
                    .Fields("REM_OrderDate").Value = Now()
'                    .Fields("REM_ActualStartDate").Value = 0
'                    .Fields("REM_ActualDoneDate").Value = 0
                    .Fields("REM_PreviousResult").Value = "none"
                    .Fields("REM_TaskStatus").Value = iTask.TaskStatus
                    .Fields("REM_ActualJobNumber").Value = "none"
                    .Fields("REM_InhibitChgs").Value = False
                    .Fields("REM_Comment").Value = "none"
                    .Fields("CAN_Description").Value = iTask.Can.Description
                    .Fields("CAN_Volume").Value = iTask.Can.WorkingVolume
                    .Fields("CAN_WorkCap").Value = iTask.Can.WorkingCapacity
                    .Fields("RCP_Number").Value = iTask.Rcp.Number
                    .Fields("RCP_Name").Value = iTask.Rcp.Name
                    .Fields("PRG_Number").Value = iTask.Rcp.Purge_ProfileNumber
'                    .Fields("PRG_Name").Value = iTask.Rcp.Purge_ProfileName
                    .Fields("SEQ_Number").Value = 0
                    .Fields("AVL_FileRoot").Value = iTask.AVL_FileRoot
                    .Update
                    DoEvents
                Else
                    ' Error - One or more Records Returned
                    errFlag = True
                    Write_ELog "RemoteTask Entry Failure - Record already exists for TaskID# " & iTask.TaskID
                End If
                   
            End With
            mobjRemTaskRst.Close
        
    frmSearchRemote.adoRemoteTasks.Refresh
    frmSearchRemote.dgRemoteTasks.Refresh
    frmSearchRemote.Refresh

'    If (USINGREMAVLFILES) Then
'        Select Case errFlag
'            Case True
'                iTask.TaskStatus = "RemoteTask Entry Failure - Record already exists for TaskID# " & iTask.TaskID
'                AVL_TaskFile_Move iTask, "Request", "Failed"
'            Case False
'                AVL_TaskFile_Move iTask, "Request", "OnList"
'        End Select
'
'End If

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

Public Sub AVL_TaskFile_Move(ByRef iTask As RemTaskControlBlock, ByVal srcFolder As String, ByVal dstFolder As String)
'
'        Move (& Update) AVL Host Task Order iTask's AVLfile from srcFolder to dstFolder
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 31592
Dim srcFilePath As String
Dim dstFilePath As String
Dim sfilePath As String
Dim dfilePath As String
Dim iSrcFileNumber As Integer
Dim iDstFileNumber As Integer
Dim sDir As String
Dim sFile As String
Dim sPath As String
Dim dDir As String
Dim dFile As String
Dim dPath As String
Dim strExists As String
Dim linestr As String
Dim linestr2 As String
Dim descstr As String
Dim valuestr As String
Dim chrnumEquals As Integer
Dim bDone As Boolean
Dim fs, f As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Dim fso As New FileSystemObject, dstFile, srcFile, tmpFile As File, ts As TextStream
        
Debug.Print "Moving AVL Task Order File (" & iTask.AVL_FileRoot & ") from " & srcFolder & " to " & dstFolder
    frmMainMenu.MousePointer = vbHourglass
    frmDelayBox.MousePointer = vbHourglass

ChgErrModule 22, 3147
    srcFilePath = FILEPATH_avltasks & srcFolder & "\" & iTask.AVL_FileRoot
    dstFilePath = FILEPATH_avltasks & dstFolder & "\" & iTask.AVL_FileRoot
ChgErrModule 22, 3148
        
'Debug.Print "FileName = " iTask.AVL_FileRoot
'    iSrcFileNumber = FreeFile
'    sfilePath = srcFilePath & iTask.AVL_FileRoot
'    Open srcFilePath For Input As #iSrcFileNumber
    Set srcFile = fso.GetFile(srcFilePath)
    Set ts = srcFile.OpenAsTextStream(ForReading)
ChgErrModule 22, 3149

'    iDstFileNumber = FreeFile
'    dfilePath = dstFilePath & iTask.AVL_FileRoot
'    Open dstFilePath For Output As #iDstFileNumber
    Set dstFile = fso.CreateTextFile(dstFilePath, True)
    bDone = False
ChgErrModule 22, 3150
    Do While ((Not bDone) And (Not ts.AtEndOfStream))
'    Do While ((Not bDone) And (Not EOF(srcFile)))
'    Do While (Not bDone)
ChgErrModule 22, 3151
'        Input #iSrcFileNumber, linestr
        linestr = ts.ReadLine

ChgErrModule 22, 3152
'Debug.Print "Line - " & linestr
        chrnumEquals = InStr(1, linestr, "=")
        descstr = Mid(linestr, 1, (chrnumEquals - 1))
        valuestr = Mid(linestr, (chrnumEquals + 1), (Len(linestr) - chrnumEquals))
'Debug.Print "Dsc - " & descstr & " <><><> " & "Val - " & valuestr
        Select Case Trim(descstr)
            Case "ActualBaseUnitAndShift"
                If ((iTask.ActualStation <> 0) And (iTask.ActualShift <> 0)) Then
                    linestr2 = descstr & "=" & STN_INFO(iTask.ActualStation).SysID & Format(iTask.ActualShift, "0")
                Else
                    linestr2 = linestr
                End If
            Case "JobNumber"
                If (Len(iTask.JobNumber) > 1) Then
                    linestr2 = descstr & "=" & iTask.JobNumber
                Else
                    linestr2 = linestr
                End If
            Case "Status"
                linestr2 = descstr & "=" & iTask.TaskStatus
                bDone = True
            Case Else
                linestr2 = linestr
        End Select
ChgErrModule 22, 3153
'        Write #iDstFileNumber, linestr2
        dstFile.WriteLine linestr2
ChgErrModule 22, 3154
    Loop
ChgErrModule 22, 3155
    If (ts.AtEndOfStream And (Not bDone)) Then
'    If (Not bDone) Then
        linestr2 = descstr & "=" & iTask.TaskStatus
        dstFile.WriteLine linestr2
    End If
    ts.Close
    dstFile.Close
    
ChgErrModule 22, 3156
    strExists = Dir(dstFilePath)
    If strExists <> "" Then
Debug.Print "File Updated: " & dstFilePath
ChgErrModule 22, 3157
        Set srcFile = fso.GetFile(srcFilePath)
        srcFile.Delete
ChgErrModule 22, 3158
Debug.Print "File Deleted: " & srcFilePath
    End If
    
ChgErrModule 22, 3159
    If (InStr(1, UCase(iTask.TaskStatus), UCase("InActiv")) > 0) Then
        srcFilePath = FILEPATH_avltasks & "Request" & "\" & iTask.AVL_FileRoot
        strExists = Dir(srcFilePath)
        If strExists <> "" Then
            Set srcFile = fso.GetFile(srcFilePath)
            srcFile.Delete
'            Kill srcFilePath
Debug.Print "ReqIna File Deleted: " & srcFilePath
        Else
            Write_ELog "Task Inactivate file not found at " & srcFilePath
Debug.Print "ReqIna File not found at " & srcFilePath
        End If
    End If
       
    frmMainMenu.MousePointer = vbDefault
    frmDelayBox.MousePointer = vbDefault
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Debug.Print "error moving a task order file"
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

Public Sub AVL_StatusFile_Update(ByVal iStn As Integer, ByVal iShift As Integer)
'
'        Update AVL Unit Status File
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 31290
Dim sFolder As String
Dim sFileRoot As String
Dim srcFileName As String
Dim dstFileName As String
Dim tmpFileName As String
Dim srcFilePath As String
Dim dstFilePath As String
Dim tmpFilePath As String
Dim iSrcFileNumber As Integer
Dim iDstFileNumber As Integer
Dim sDir As String
Dim sFile As String
Dim sPath As String
Dim dDir As String
Dim dFile As String
Dim dPath As String
Dim strExists As String
Dim linestr As String
Dim linestr2 As String
Dim descstr As String
Dim valuestr As String
Dim chrnumEquals As Integer
Dim srcFlag As Boolean
Dim bDone As Boolean
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
Dim fso As New FileSystemObject, dstFile, srcFile, tmpFile As File, ts As TextStream

        
    sFolder = FILEPATH_avlstatus
    sFileRoot = STN_INFO(iStn).SysID & Format(iShift, "0")
    sFile = sFolder & sFileRoot & ".txt"
    
ChgErrModule 22, 31291
    strExists = Dir(sFile)
    srcFlag = IIf((strExists <> ""), True, False)
    If srcFlag Then
        ' a file already exists
'Debug.Print "Found Source File: " & strExists
        srcFileName = strExists
        dstFileName = Mid(srcFileName, 1, (Len(srcFileName) - 3)) & "tmp"
        srcFilePath = sFolder & srcFileName
        
    Else
        ' no existing file
        srcFileName = sFileRoot & ".txt"
        dstFileName = sFileRoot & ".tmp"
    End If
    
    dstFilePath = sFolder & dstFileName
    
ChgErrModule 22, 31292
'    iDstFileNumber = FreeFile
'    Open dstFilePath For Output As #iDstFileNumber
    Set dstFile = fso.CreateTextFile(dstFilePath, True)
'    Set ts = dstFile.OpenAsTextStream(ForWriting)
    bDone = False
    
    ' Mode
    descstr = "Mode"
    valuestr = ModeDescLong(StationControl(iStn, iShift).Mode)
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' Phase
    descstr = "Phase"
    Select Case StationControl(iStn, iShift).Mode
        Case VBLOAD
            valuestr = LoadPhaseDesc(LoadControl(iStn, iShift).Phase)
        Case VBPURGE
            valuestr = PurgePhaseDesc(PurgeControl(iStn, iShift).Phase)
        Case Else
            valuestr = "na"
    End Select
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' Cycle
    descstr = "Cycle"
    valuestr = Format(StationControl(iStn, iShift).CurrCycle, "###0")
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' JobNumber
    descstr = "JobNumber"
    valuestr = StationControl(iStn, iShift).Job_Number
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' TaskID
    descstr = "TaskID"
    valuestr = StnRemoteTask(iStn, iShift).TaskID
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' VIN
    descstr = "VINNumber"
    valuestr = StnRemoteTask(iStn, iShift).VIN
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    ' ALARM MSG
    descstr = "Alarm"
    If (StationControl(iStn, iShift).Mode = VBPAUSEALARM) Then
        valuestr = StationControl(iStn, iShift).PauseMessage
    Else
        valuestr = " "
    End If
    linestr2 = descstr & "=" & valuestr
'    Write #iDstFileNumber, linestr2
    dstFile.WriteLine linestr2
    
    
'    Close #iDstFileNumber
    dstFile.Close
    
ChgErrModule 22, 31293
    
    
    
    strExists = Dir(dstFilePath)
    If (strExists <> "") Then
        tmpFilePath = dstFilePath
ChgErrModule 22, 31294
'Debug.Print "File Updated: " & dstFilePath
        If srcFlag Then
ChgErrModule 22, 31295
            strExists = Dir(srcFilePath)
            If (strExists <> "") Then
ChgErrModule 22, 31296
                Set srcFile = fso.GetFile(srcFilePath)
'                Kill srcFilePath
                srcFile.Delete
'Debug.Print "File Deleted: " & srcFilePath
            End If
        End If
                ' Moving file
                ' Get a handle to the tmp file
ChgErrModule 22, 31297
                Set tmpFile = fso.GetFile(tmpFilePath)
                ' Move the file to dst file
                dstFileName = Mid(srcFileName, 1, (Len(srcFileName) - 3)) & "txt"
                dstFilePath = sFolder & srcFileName
ChgErrModule 22, 31298
                tmpFile.Move (dstFilePath)
'                dstFile.Close
'Debug.Print dstFilePath & " Moved to " & srcFilePath
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

Sub RemStatus_Update(iStn As Integer, iShift As Integer)
' Procedure Name:   RemStatus_Update
' Created by:       MMW
' Description:      This routine updates the StationStatus table
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 22, 442
Dim doFlag As Boolean
Dim Criteria As String

    ' has anything changed ??
    doFlag = True
'    doFlag = False
'    If (StationRemStatusControl(iStn, iShift).Mode_LastStatus <> StationControl(iStn, iShift).Mode) Then doFlag = True
'    Select Case StationControl(iStn, iShift).Mode
'        Case VBLOAD
'            If (StationRemStatusControl(iStn, iShift).Phase_LastStatus <> LoadControl(iStn, iShift).Phase) Then doFlag = True
'        Case VBPURGE
'            If (StationRemStatusControl(iStn, iShift).Phase_LastStatus <> PurgeControl(iStn, iShift).Phase) Then doFlag = True
'        Case Else
'            If (StationRemStatusControl(iStn, iShift).Phase_LastStatus <> 0) Then doFlag = True
'    End Select
'    If (StationRemStatusControl(iStn, iShift).Cycle_LastStatus <> StationControl(iStn, iShift).CurrCycle) Then doFlag = True
    
ChgErrModule 22, 1440
    ' update (if required)
    If doFlag Then

            ' Open StationStatus Records
            Criteria = "SELECT * FROM [StationStatus] WHERE ([StationStatus].[Station] = " & iStn & " AND [StationStatus].[Shift] = " & iShift & " ) "
            mobjRemStatusRst.CursorLocation = adUseClient
            mobjRemStatusRst.Open Criteria, mobjRemStatusConn, adOpenDynamic, adLockOptimistic, adCmdText
            
ChgErrModule 22, 1441
            With mobjRemStatusRst
            
                If .BOF Then
ChgErrModule 22, 1442
                    .AddNew
                    .Fields("Station").Value = iStn
                    .Fields("Shift").Value = iShift
                    If (USINGREMCANLOAD And (Len(StnRemoteTask(iStn, iShift).TaskID) > 8)) Then
                        .Fields("TaskID").Value = StnRemoteTask(iStn, iShift).TaskID
                    Else
                        .Fields("TaskID").Value = "na"
                    End If
                    .Fields("VIN").Value = JobInfo(iStn, iShift).Vehicle
                Else
ChgErrModule 22, 1443
                    .MoveFirst
'                    .Open
                End If
ChgErrModule 22, 1444
                    .Fields("Mode").Value = ModeDescLong(StationControl(iStn, iShift).Mode)
                    Select Case StationControl(iStn, iShift).Mode
                        Case VBLOAD
                            .Fields("Phase").Value = LoadPhaseDesc(LoadControl(iStn, iShift).Phase)
                        Case VBPURGE
                            .Fields("Phase").Value = PurgePhaseDesc(PurgeControl(iStn, iShift).Phase)
                        Case Else
                            .Fields("Phase").Value = "na"
                    End Select
                    .Fields("Cycle").Value = StationControl(iStn, iShift).CurrCycle
                    .Update
                    DoEvents
ChgErrModule 22, 1445
                   
            End With
            mobjRemStatusRst.Close
            
' Debug.Print "Station #" & Format(iStn, "#0") & "/Shift #" & Format(iShift, "0") & " RemStatus_Update @ " & Format(Now, "YYYY MMM DD  hh:mm:ss") & " - " & Format(Timer, "###,##0.000")
            
ChgErrModule 22, 1446
            ' update "Last" values
            StationRemStatusControl(iStn, iShift).Mode_LastStatus = StationControl(iStn, iShift).Mode
            Select Case StationControl(iStn, iShift).Mode
                Case VBLOAD
                    StationRemStatusControl(iStn, iShift).Phase_LastStatus = LoadControl(iStn, iShift).Phase
                Case VBPURGE
                    StationRemStatusControl(iStn, iShift).Phase_LastStatus = PurgeControl(iStn, iShift).Phase
                Case Else
                    StationRemStatusControl(iStn, iShift).Phase_LastStatus = 0
            End Select
            StationRemStatusControl(iStn, iShift).Cycle_LastStatus = StationControl(iStn, iShift).CurrCycle
            
ChgErrModule 22, 1447
           If (USINGREMAVLFILES) Then AVL_StatusFile_Update iStn, iShift
        
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

Sub Manip_Files()
    Dim fso As New FileSystemObject, txtfile, fil1, fil2
    Set txtfile = fso.CreateTextFile("c:\testfile.txt", True)
    MsgBox "Writing file"
    ' Write a line.
    txtfile.Write ("This is a test.")
    ' Close the file to writing.
    txtfile.Close
    MsgBox "Moving file to c:\tmp"
    ' Get a handle to the file in root of C:\.
    Set fil1 = fso.GetFile("c:\testfile.txt")
    ' Move the file to \tmp directory.
    fil1.Move ("c:\tmp\testfile.txt")
    MsgBox "Copying file to c:\temp"
    ' Copy the file to \temp.
    fil1.Copy ("c:\temp\testfile.txt")
    MsgBox "Deleting files"
    ' Get handles to files' current location.
    Set fil1 = fso.GetFile("c:\tmp\testfile.txt")
    Set fil2 = fso.GetFile("c:\temp\testfile.txt")
    ' Delete the files.
    fil1.Delete
    fil2.Delete
    MsgBox "All done!"
End Sub

