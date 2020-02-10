Attribute VB_Name = "Module20"
' Module 20  - Task Order Manager Logic
'
Option Explicit

Public Function RemData_Clear() As RemData
'
Dim xTOM As RemData
    xTOM.TaskID = "na"
    xTOM.VIN = "na"
    xTOM.TaskType = "na"
    xTOM.ActualShift = 0
    xTOM.ActualStation = 0
    xTOM.CanVolume = 0
    xTOM.canWC = 0
    xTOM.rcpRecipeNumber = 0
    xTOM.rcpUse2Gm = False
    xTOM.rcpUseWCM = False
    xTOM.RequestedShift = 0
    xTOM.RequestedStation = 0
    xTOM.TaskStatus = "na"
    xTOM.PreviousResult = "na"
    xTOM.ReportComplete = False
    
    RemData_Clear = xTOM
End Function

Public Function ValidTomRecipe(thisRcp As Integer) As Boolean
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
    ValidTomRecipe = IIf(errorFlag, False, True)
End Function

Sub TomTask_Update(stn As Integer, Shift As Integer, newStatus As String, prevResult As String)
' Procedure Name:   TomTask_Update
' Created by:       MMW
' Description:      This routine updates the TomTask status
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 20, 44

Dim Criterion, filename As String
Dim dbDbase As Database
Dim rsTable As Recordset
Dim msg As String

    ' update TOM Task status in station/shift
    StnRemoteTask(stn, Shift).TaskStatus = newStatus
    ' update TOM Task status in DB
    Criterion = _
        "SELECT * FROM [TOM_CanLoadTasks] WHERE [TOM_CanLoadTasks].[TOM_TestOrderID] = '" & _
         StnRemoteTask(stn, Shift).TaskID & "'"
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATAREM)
    Set rsTable = dbDbase.OpenRecordset(Criterion, dbOpenDynaset)
    
    If rsTable.BOF Then
    
        ' no record found
        msg = "TOM Task Status Update Failed for TaskID " & StnRemoteTask(stn, Shift).TaskID
        msg = msg & "  , Station " & Format(stn, "0")
        msg = msg & "  , Shift " & Format(Shift, "0")
        msg = msg & "  , - No record found)"
        Write_ELog msg
        
    Else
    
        ' record(s) found
        rsTable.MoveFirst
        If (rsTable.RecordCount > 1) Then
        
            ' more than one record found
            msg = "TOM Task Status Update Failed for TaskID " & StnRemoteTask(stn, Shift).TaskID
            msg = msg & "  , Station " & Format(stn, "0")
            msg = msg & "  , Shift " & Format(Shift, "0")
            msg = msg & "  , - More than one record found)"
            Write_ELog msg
            
        Else
        
            ' one record found
            rsTable.Edit
            Select Case newStatus
                Case "Ready"
                    rsTable("TOM_TaskStatus") = newStatus
                    rsTable("TOM_PreviousResult") = prevResult
                    rsTable("TOM_ActualStartDate") = 0
                    rsTable("TOM_ActualDoneDate") = 0
                    StnRemoteTask(stn, Shift).PreviousResult = prevResult
                Case "Active"
                    rsTable("TOM_TaskStatus") = newStatus
                    rsTable("TOM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
                    rsTable("TOM_ActualStation") = stn
                    rsTable("TOM_ActualShift") = Shift
                    rsTable("TOM_ActualStartDate") = Now()
                Case "Done"
                    rsTable("TOM_TaskStatus") = newStatus
                    rsTable("TOM_ActualJobNumber") = Format(StationControl(stn, Shift).Job_Number, "000000")
                    rsTable("TOM_ActualStation") = stn
                    rsTable("TOM_ActualShift") = Shift
                    rsTable("TOM_ActualDoneDate") = Now()
                Case "Invalid"
                    rsTable("TOM_TaskStatus") = newStatus
                Case Else
            End Select
            
        End If
        
    End If
    
    
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



