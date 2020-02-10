Attribute VB_Name = "Module20"
' module 20      WaterBath Heater/Chiller Temperature Control
'
Option Explicit


Public Sub ChillerCommander()
Dim sMsg As String
    ' initialization
    If Not LF_Chiller.InitComplete Then
        ChillerInit
    End If
    If Not LF_Chiller.CommOnline Then
        Exit Sub
    End If
    If Not LF_Chiller.InitComplete Then
        frmMainMenu.ChillerCommInit
    End If
    If Not LF_Chiller.CommOK Then
        LF_Chiller.PvIn = DegFtoC(-99#)
        Exit Sub
    End If
    If Not LF_Chiller.RunChiller Then
        Exit Sub
    End If
            
    If LF_Chiller.ChillerPhase = 0 Then ChillerPortWrite chillerIn_PV, 0
    If LF_Chiller.ChillerPhase = 0 Then LF_Chiller.ChillerPhase = 1
    LF_Chiller.SpOut = IIf(USINGC, WaterBathSP, DegFtoC(CSng(WaterBathSP)))
    ' has a command completed? or timed-out?
    If LF_Chiller.CurCmdComplete Then
        ' Current Command Completed Successfully
        ChillerResponse(0) = LF_Chiller.CmdRecChars
        If LF_Chiller.CurCmdIdx = chillerOut_Start Then LF_Chiller.ChillerRunning = True
        If LF_Chiller.CurCmdIdx = chillerOut_Stop Then LF_Chiller.ChillerRunning = False
        LF_Chiller.ChillerPhase = IIf(LF_Chiller.ChillerPhase < 20, LF_Chiller.ChillerPhase + 1, 1)
        ' send next command
        If LF_Chiller.CommOK Then ChillerCommandSelect
    ElseIf LF_Chiller.CmdTimeoutTimer < Timer Then
        ' Current Command Timed Out (i.e. failed)
        Write_ELog "Chiller Comm Timeout"
        ChillerResponse(0) = "timeout"
        LF_Chiller.CurCmdTimeout = True
        LF_Chiller.ErrorCount = LF_Chiller.ErrorCount + 1
        LF_Chiller.CommOK = IIf(LF_Chiller.ErrorCount > LF_Chiller.MaxErrorCount, False, True)
        ' repeat same command
        If LF_Chiller.CommOK Then ChillerCommandSelect
    ElseIf LF_Chiller.CmdRecErrorFlag Then
        ' Current Command returned an error
        Write_ELog "Chiller Comm Error >" & LF_Chiller.CmdRecChars & "<"
        ChillerResponse(0) = LF_Chiller.CmdRecChars
        LF_Chiller.ErrorCount = LF_Chiller.ErrorCount + 1
        LF_Chiller.CommOK = IIf(LF_Chiller.ErrorCount > LF_Chiller.MaxErrorCount, False, True)
        ' repeat same command
        If LF_Chiller.CommOK Then ChillerCommandSelect
    End If
    ' check for loss of CommOK
    If Not LF_Chiller.CommOK Then
        ' always log to the Event Log
        Write_ELog "Heater Comm NOT OK"
    End If
End Sub

Public Sub ChillerCommandSelect()
    Select Case LF_Chiller.ChillerPhase
        Case 1
            ' read PV
            ChillerPortWrite chillerIn_PV, 0
        Case 2
            ' read SP
            ChillerPortWrite chillerIn_SP, 0
        Case 3
            ' write SP
            If LF_Chiller.SpIn <> LF_Chiller.SpOut And LF_Chiller.ChillerRunning Then
                ' write SP
                ChillerPortWrite chillerOut_SP, LF_Chiller.SpOut
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 4
            ' read pump output
            ChillerPortWrite chillerIn_Out, 0
        Case 5
            ' write pump output
'            LF_Chiller.OutOut = LF_Chiller.OutIn
            If LF_Chiller.OutIn <> LF_Chiller.OutOut And LF_Chiller.ChillerRunning Then
                ' write output
                ChillerPortWrite chillerOut_Out, LF_Chiller.OutOut
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 6
            ' read Operating Mode
            ChillerPortWrite chillerIn_OperMode, 0
        Case 7
            ' write Operating Mode
            If LF_Chiller.OperModeIn <> LF_Chiller.OperModeOut And LF_Chiller.ChillerRunning Then
                ' write Operating Mode
                ChillerPortWrite chillerOut_OperMode, LF_Chiller.OperModeOut
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 8
            ' read Overtemp SetPoint
            ChillerPortWrite chillerIn_OvrTmpSp, 0
        Case 9
            ' read P
            ChillerPortWrite chillerIn_P, 0
        Case 10
            ' write P
            LF_Chiller.P_Out = LF_Chiller.P_In
            If LF_Chiller.P_In <> LF_Chiller.P_Out Then
                ' write P
                ChillerPortWrite chillerOut_P, LF_Chiller.P_Out
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 11
            ' read I
            ChillerPortWrite chillerIn_I, 0
        Case 12
            ' write I
            LF_Chiller.I_Out = LF_Chiller.I_In
            If LF_Chiller.I_In <> LF_Chiller.I_Out Then
                ' write I
                ChillerPortWrite chillerOut_I, LF_Chiller.I_Out
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 13
            ' read Mode
            ChillerPortWrite chillerIn_Mode, 0
        Case 14
            ' write Mode
            If LF_Chiller.ModeIn <> LF_Chiller.ModeOut Then
                ' write Mode
                ChillerPortWrite chillerOut_Mode, LF_Chiller.ModeOut
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 15
            ' read Chiller Type
            ChillerPortWrite chillerIn_Type, 0
        Case 16
            ' read Software Version
            ChillerPortWrite chillerIn_Version, 0
        Case 17
            ' read PV
            ChillerPortWrite chillerIn_PV, 0
        Case 18
            ' read Status
            ChillerPortWrite chillerIn_Status, 0
        Case 19
            ' write START or STOP
            If Not LF_Chiller.RunChiller And LF_Chiller.ChillerRunning Then
                ' write STOP
                ChillerPortWrite chillerOut_Stop, 0
            ElseIf Not LF_Chiller.ChillerRunning And LF_Chiller.RunChiller Then
                ' write START
                ChillerPortWrite chillerOut_Start, 0
            Else
                ' read PV
                ChillerPortWrite chillerIn_PV, 0
            End If
        Case 20
            ' read Stat bits
            ChillerPortWrite chillerIn_Stat, 0
                  
    End Select
End Sub

Public Sub ChillerStart()
Dim LoopCount As Long
    LoopCount = 0
    ' turn on "run chiller"
    LF_Chiller.RunChiller = True
    ' wait for Chiller to Start
    Do While (Not LF_Chiller.ChillerRunning And LoopCount < 100000)
        DoEvents
        LoopCount = LoopCount + 1
    Loop
End Sub

Public Sub ChillerStop()
Dim LoopCount As Long
    LoopCount = 0
    ' turn off "run chiller"
    LF_Chiller.RunChiller = False
    ' wait for Chiller to Stop
    Do While (LF_Chiller.ChillerRunning And LoopCount < 100000)
        DoEvents
        LoopCount = LoopCount + 1
    Loop
End Sub

Public Sub ChillerInit()
Dim tmpTimer As Single
    LF_Chiller.RunChiller = True
'    LF_Chiller.RunChiller = False
    LF_Chiller.BufferIn = ""
    LF_Chiller.BufferOut = ""
    LF_Chiller.ChillerPhase = 0
    LF_Chiller.ChillerRunning = False
    LF_Chiller.CmdRecAckFlag = False
    LF_Chiller.CmdRecChars = ""
    LF_Chiller.CmdRecErrorFlag = False
    LF_Chiller.CmdRecErrorNumber = 0
    LF_Chiller.CmdRecValueChars = ""
    LF_Chiller.CmdSentFlag = False
    LF_Chiller.CmdToBeAckFlag = False
    LF_Chiller.CmdToBeSentFlag = False
    LF_Chiller.CommOK = True
    LF_Chiller.CommOnline = True
    LF_Chiller.CurCmdChars = ""
    LF_Chiller.CurCmdComplete = False
    LF_Chiller.CurCmdDesc = "none"
    LF_Chiller.CurCmdIdx = 0
    LF_Chiller.CurCmdTimeout = False
    LF_Chiller.ErrorCount = 0
    LF_Chiller.I_In = 0
    LF_Chiller.I_Out = 0
    LF_Chiller.InitComplete = False
    LF_Chiller.IntFaultMc1 = False
    LF_Chiller.IntFaultMc2 = False
    LF_Chiller.LowLevel = False
    LF_Chiller.MaxErrorCount = 3
    LF_Chiller.ModeIn = 0
    LF_Chiller.ModeOut = 0
    LF_Chiller.OperModeIn = 0
'    LF_Chiller.OperModeOut = 3
    LF_Chiller.OperModeOut = 2
    LF_Chiller.OutIn = 0
'    LF_Chiller.OutOut = 0
    LF_Chiller.OutOut = pumpNORMAL
    LF_Chiller.Overtemp = False
    LF_Chiller.OvrTmpSpIn = 0
    LF_Chiller.P_In = 0
    LF_Chiller.P_Out = 0
    LF_Chiller.PumpBlocked = False
    LF_Chiller.PvIn = 0
    LF_Chiller.SpIn = 0
    LF_Chiller.SpOut = IIf(USINGC, Com_AIO(acAmbTempSensor).EUValue, DegFtoC(Com_AIO(acAmbTempSensor).EUValue))            ' ambient temp
    LF_Chiller.StatIn = "00000"
    LF_Chiller.StatusIn = 0
'    LF_Chiller.Type = ?
'    LF_Chiller.Version = ?
    LF_Chiller.TimeoutValue = Chiller_Timeout
    tmpTimer = Timer + LF_Chiller.TimeoutValue
    If tmpTimer > 86400 Then tmpTimer = tmpTimer - 86400
    LF_Chiller.CmdTimeoutTimer = tmpTimer
End Sub

Public Function ChillerOK() As Boolean
Dim tmpProblem As Boolean
    tmpProblem = False
    If LF_Chiller.Overtemp <> False Then tmpProblem = True
    If LF_Chiller.LowLevel <> False Then tmpProblem = True
    If LF_Chiller.PumpBlocked <> False Then tmpProblem = True
    If LF_Chiller.IntFaultMc1 <> False Then tmpProblem = True
    If LF_Chiller.IntFaultMc2 <> False Then tmpProblem = True
    If LF_Chiller.StatusIn <> 0 Then tmpProblem = True
    If LF_Chiller.CommOK <> True Then tmpProblem = True
    If LF_Chiller.CommOnline <> True Then tmpProblem = True
    ChillerOK = IIf(tmpProblem, False, True)
End Function

Public Function ChillerReady() As Boolean
Dim tmpProblem As Boolean
    tmpProblem = False
    If LF_Chiller.Type <> Null Then tmpProblem = True
    If LF_Chiller.Version <> Null Then tmpProblem = True
    If LF_Chiller.PvIn <> Null Then tmpProblem = True
    If LF_Chiller.SpIn <> Null Then tmpProblem = True
    If LF_Chiller.ModeIn <> 0 Then tmpProblem = True
    If LF_Chiller.OperModeIn <> False Then tmpProblem = True
    If LF_Chiller.Overtemp <> False Then tmpProblem = True
    ChillerReady = IIf(tmpProblem, False, True)
End Function

Public Sub ChillerRun()
Dim tmpTimer As Single
'    PID_INFO(tcbWaterBathTemp).timeOut = False
    tmpTimer = Timer + LF_Chiller.TimeoutValue
    If tmpTimer > 86400 Then tmpTimer = tmpTimer - 86400
    LF_Chiller.CmdTimeoutTimer = tmpTimer
    LF_Chiller.CommOK = True
    LF_Chiller.CommOnline = True
    LF_Chiller.ErrorCount = 0
    LF_Chiller.RunChiller = True
    If LF_Chiller.ChillerPhase <> 0 Then ChillerPortWrite chillerIn_PV, 0
    Write_ELog "Chiller Comm OK"
End Sub

Public Sub ChillerPortWrite(ByVal cmdIdx As Integer, ByVal cmdVal As Single)
Dim tmpTimer As Single
    LF_Chiller.CurCmdIdx = cmdIdx
    Select Case LF_Chiller.CurCmdIdx
        Case chillerIn_PV
            LF_Chiller.CurCmdDesc = "Read PV"
            LF_Chiller.CurCmdChars = "IN_PV_00"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_SP
            LF_Chiller.CurCmdDesc = "Read SP"
            LF_Chiller.CurCmdChars = "IN_SP_00"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Out
            LF_Chiller.CurCmdDesc = "Read Output"
            LF_Chiller.CurCmdChars = "IN_SP_01"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_OperMode
            LF_Chiller.CurCmdDesc = "Read Oper Mode"
            LF_Chiller.CurCmdChars = "IN_SP_02"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_OvrTmpSp
            LF_Chiller.CurCmdDesc = "Read OverTemp SP"
            LF_Chiller.CurCmdChars = "IN_SP_03"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_P
            LF_Chiller.CurCmdDesc = "Read P"
            LF_Chiller.CurCmdChars = "IN_PAR_00"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_I
            LF_Chiller.CurCmdDesc = "Read I"
            LF_Chiller.CurCmdChars = "IN_PAR_01"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Mode
            LF_Chiller.CurCmdDesc = "Read Mode"
            LF_Chiller.CurCmdChars = "IN_MODE_00"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Type
            LF_Chiller.CurCmdDesc = "Read Type"
            LF_Chiller.CurCmdChars = "TYPE"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Version
            LF_Chiller.CurCmdDesc = "Read Version"
            LF_Chiller.CurCmdChars = "VERSION_R"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Status
            LF_Chiller.CurCmdDesc = "Read Status"
            LF_Chiller.CurCmdChars = "STATUS"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerIn_Stat
            LF_Chiller.CurCmdDesc = "Read Stat Bits"
            LF_Chiller.CurCmdChars = "STAT"
            LF_Chiller.CmdToBeAckFlag = False
        Case chillerOut_SP
            LF_Chiller.CurCmdDesc = "Write SP"
            LF_Chiller.CurCmdChars = "OUT_SP_00_" & Format(cmdVal, "##0.00")
            LF_Chiller.SpOut = cmdVal
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_Out
            LF_Chiller.CurCmdDesc = "Write Output"
            LF_Chiller.CurCmdChars = "OUT_SP_01_" & Format(cmdVal, "##0")
            LF_Chiller.OutOut = CInt(cmdVal)
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_OperMode
            LF_Chiller.CurCmdDesc = "Write Oper Mode"
            LF_Chiller.CurCmdChars = "OUT_SP_02_" & Format(cmdVal, "##0")
            LF_Chiller.OperModeOut = CInt(cmdVal)
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_P
            LF_Chiller.CurCmdDesc = "Write P"
            LF_Chiller.CurCmdChars = "OUT_PAR_00_" & Format(cmdVal, "##0.00")
            LF_Chiller.P_Out = cmdVal
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_I
            LF_Chiller.CurCmdDesc = "Write I"
            LF_Chiller.CurCmdChars = "OUT_PAR_01_" & Format(cmdVal, "##0")
            LF_Chiller.I_Out = cmdVal
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_Mode
            LF_Chiller.CurCmdDesc = "Write Mode"
            LF_Chiller.CurCmdChars = "OUT_MODE_00_" & Format(cmdVal, "0")
            LF_Chiller.ModeOut = CInt(cmdVal)
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_Start
            LF_Chiller.CurCmdDesc = "Start Chiller"
            LF_Chiller.CurCmdChars = "START"
            LF_Chiller.CmdToBeAckFlag = True
        Case chillerOut_Stop
            LF_Chiller.CurCmdDesc = "Stop Chiller"
            LF_Chiller.CurCmdChars = "STOP"
            LF_Chiller.CmdToBeAckFlag = True
    End Select
    ChillerShuffleCommandHistory LF_Chiller.CurCmdChars
    LF_Chiller.CurCmdIdx = cmdIdx
    LF_Chiller.BufferOut = LF_Chiller.CurCmdChars & vbCrLf
    tmpTimer = (0.001 * LF_Chiller.TimeoutValue)
    tmpTimer = tmpTimer + Timer
    If tmpTimer > 86400 Then tmpTimer = tmpTimer - 86400
    LF_Chiller.CmdTimeoutTimer = tmpTimer
    LF_Chiller.CurCmdTimeout = False
    LF_Chiller.CurCmdComplete = False
    LF_Chiller.CmdRecAckFlag = False
    LF_Chiller.CmdRecErrorFlag = False
    LF_Chiller.CmdRecErrorNumber = 0
    LF_Chiller.CmdRecChars = ""
    LF_Chiller.CmdRecValueChars = ""
    LF_Chiller.CmdSentFlag = False
    LF_Chiller.CmdToBeSentFlag = True
    frmMainMenu.MSComm(mscommChiller).Output = LF_Chiller.BufferOut
End Sub

Public Sub ChillerShuffleCommandHistory(ByVal newCommand As String)
    ChillerCommands(7) = ChillerCommands(6)
    ChillerCommands(6) = ChillerCommands(5)
    ChillerCommands(5) = ChillerCommands(4)
    ChillerCommands(4) = ChillerCommands(3)
    ChillerCommands(3) = ChillerCommands(2)
    ChillerCommands(2) = ChillerCommands(1)
    ChillerCommands(1) = ChillerCommands(0)
    ChillerCommands(0) = newCommand
    ChillerResponse(7) = ChillerResponse(6)
    ChillerResponse(6) = ChillerResponse(5)
    ChillerResponse(5) = ChillerResponse(4)
    ChillerResponse(4) = ChillerResponse(3)
    ChillerResponse(3) = ChillerResponse(2)
    ChillerResponse(2) = ChillerResponse(1)
    ChillerResponse(1) = ChillerResponse(0)
    ChillerResponse(0) = ""
End Sub




