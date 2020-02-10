VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAKServer 
   Caption         =   "AK Socket Server"
   ClientHeight    =   11175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10815
   ControlBox      =   0   'False
   Icon            =   "frmAKServer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11175
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   4575
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Text            =   "frmAKServer.frx":57E2
      Top             =   6000
      Width           =   10575
   End
   Begin VB.TextBox txtRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   8805
      TabIndex        =   8
      Text            =   "123"
      Top             =   10800
      Width           =   495
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "&Reset Server"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   10680
      Width           =   1215
   End
   Begin VB.TextBox txtCompleted 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6855
      TabIndex        =   5
      Text            =   "123456789"
      Top             =   10800
      Width           =   975
   End
   Begin VB.CommandButton cmdCfg 
      DisabledPicture =   "frmAKServer.frx":57EC
      DownPicture     =   "frmAKServer.frx":5EEE
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   10240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAKServer.frx":65F0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   700
      UseMaskColor    =   -1  'True
      Width           =   455
   End
   Begin VB.CommandButton cmdExit 
      DisabledPicture =   "frmAKServer.frx":6CF2
      DownPicture     =   "frmAKServer.frx":73F4
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   455
      Left            =   10240
      Picture         =   "frmAKServer.frx":7AF6
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   455
   End
   Begin VB.Timer tmrScreen 
      Interval        =   500
      Left            =   3840
      Top             =   10680
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1035
      Left            =   1305
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmAKServer.frx":81F8
      Top             =   120
      Width           =   8805
   End
   Begin VB.PictureBox picAPS 
      Height          =   1035
      Left            =   120
      Picture         =   "frmAKServer.frx":8200
      ScaleHeight     =   975
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   120
      Width           =   1065
   End
   Begin MSWinsockLib.Winsock sockServer 
      Left            =   2160
      Top             =   10680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4575
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmAKServer.frx":B60E
      Top             =   1320
      Width           =   10575
   End
   Begin VB.Label Label2 
      Caption         =   "commands/second"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   9300
      TabIndex        =   10
      Top             =   10800
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "commands at"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   7830
      TabIndex        =   9
      Top             =   10800
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Completed"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   10800
      Width           =   855
   End
End
Attribute VB_Name = "frmAKServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public KnownAKcommands As String
Public startChar As String
Public endChar As String
Public anyChar As String
Public sepChar As String
Public portNumStr As String
Public rateStartDTS, LastCmdDTS As Date
Public cmdCounter, deltaSec, errCounter As Double
Public cmdRate As Single
Public maxCmdIdx, secADUF, secASTZ As Integer
' *************************************************************************
'
'note sockServer states
'
' value    name                 description
'  0    sckClosed               connection closed
'  1    sckOpen                 open
'  2    sckListening            listening for incoming connections
'  3    sckConnectionPending    connection pending
'  4    sckResolvingHost        resolving remote host name
'  5    sckHostResolved         remote host name successfully resolved
'  6    sckConnecting           connecting to remote host
'  7    sckConnected            connected to remote host
'  8    sckClosing              Connection Is closing
'  9    sckError                error occured
'
' *************************************************************************

Private Sub cmdCfg_Click()
    frmAKcfg.Show
End Sub

Private Sub cmdDisconnect_Click()
    txtStatus(0).text = NowPrefixString & "Disconnect by Operator" _
        & vbCrLf _
        & txtStatus(0).text
    DisconnectAK
    Write_ELog "AK Comm Link Disconnected by Operator"
End Sub

Private Sub cmdExit_Click()
    Xit
End Sub

Private Sub Form_Load()
    startChar = Chr(2)
    endChar = Chr(3)
    anyChar = AK_anychar
    sepChar = AK_sepchar
    portNumStr = AK_portNumStr
    cmdCounter = 0
    errCounter = 0
    rateStartDTS = Now
    LastCmdDTS = Now
    
    InitAKserver
    
    txtMsg.ForeColor = SaddleBrown
    txtStatus(0).ForeColor = Teal
    txtStatus(0).text = ""
    txtStatus(1).ForeColor = Goldenrod
    txtStatus(1).text = ""
    StartListening
    cmdCfg.Enabled = True
End Sub

Private Sub DisconnectAK()
    sockServer.Close
    StartListening
    cmdCfg.Enabled = True
End Sub

Private Sub StartListening()
    sockServer.LocalPort = portNumStr
    sockServer.Listen
    txtStatus(0).text = NowPrefixString & "Begin Listening" _
        & vbCrLf _
        & txtStatus(0).text
End Sub

Private Sub Xit()
    Select Case LocalPagControl.Type
        Case pagClient
            'using AK Server
            frmAKServer.Hide
        Case pagMaster
            'using AK Server
            frmAKServer.Hide
        Case pagNone, pagAlone
            ' no AK server
            Unload frmAKServer
            Set frmAKServer = Nothing
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Xit
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub picAPS_Click()
    cmdCounter = 0
    errCounter = 0
    rateStartDTS = Now
End Sub

Private Sub sockServer_ConnectionRequest(ByVal requestID As Long)
    If sockServer.State <> sckClosed Then
        sockServer.Close
    End If
    sockServer.Accept requestID
    txtStatus(0).text = NowPrefixString & "Accepted Connection from: " & sockServer.RemoteHostIP _
        & vbCrLf _
        & txtStatus(0).text
    cmdCfg.Enabled = False
    LastCmdDTS = Now
    Write_ELog "AK Comm Link Accepted Connection from: " & sockServer.RemoteHostIP
End Sub

Private Sub sockServer_DataArrival(ByVal bytesTotal As Long)
    Dim strCmd, cmdDesc, errStatStr, strDTS As String
    Dim cmdCode As String * 4
    Dim Idx, nullChar As Integer
    
    ' log the time
    LastCmdDTS = Now
    ' read the port buffer
    sockServer.GetData strCmd, vbString
    ' remove any Nulls from the end of the string
    While (InStr(strCmd, Chr$(0)) <> 0)
        nullChar = InStr(strCmd, Chr$(0))
        strCmd = Trim(Left(strCmd, nullChar - 1))
    Wend
    cmdDesc = CommandDesc(Trim(strCmd))
'    cmdDesc = "AK"
    If cmdDesc = "AK" Then
        ' command is an AK command; parse it
        Cmd_Parse True, strCmd
        ' respond to the command
        Cmd_Responder
        Select Case AKcmd_Current.cmdCode
            Case "ACFG", "ASTZ", "ATEM"
                Idx = 1
            Case Else
                Idx = 0
        End Select
        ' display command & response on AK Server screen
        txtStatus(Idx).text = DtsPrefixString & AKcmd_Current.Printable _
            & "    ~~>~>    " _
            & AKrsp_Current.Printable _
            & vbCrLf _
            & txtStatus(Idx).text
        ' increment the "completed commands" counter
        cmdCounter = cmdCounter + 1
    Else
        ' command is not an AK command; get ready to respond to it
        Cmd_Parse False, strCmd
        ' respond to the command
        Cmd_Responder
        ' display command & response on AK Server screen
        txtStatus(0).text = DtsPrefixString & AKcmd_Current.Printable _
            & "   " _
            & cmdDesc _
            & vbCrLf _
            & txtStatus(0).text
        ' write command & response to event log
        Write_ELog DtsPrefixString & AKcmd_Current.Printable & "   " & cmdDesc
        ' increment the "errors" counter
        errCounter = errCounter + 1
    End If

End Sub

Private Sub tmrScreen_Timer()
    txtMsg.text = vbCrLf & Format(Now(), "YYYY MMM D   hh:mm:ss") & vbCrLf & WinsockStateDesc(sockServer.State)
    If (Len(txtStatus(0).text) > 32000) Then txtStatus(0).text = Mid(txtStatus(0).text, 1, 24000)
    If (Len(txtStatus(1).text) > 32000) Then txtStatus(1).text = Mid(txtStatus(1).text, 1, 24000)
    txtCompleted.text = Format(cmdCounter, "###,###,##0")
    deltaSec = CDbl(DateDiff("s", rateStartDTS, Now))
    If deltaSec > 0 Then cmdRate = CSng(cmdCounter) / CSng(deltaSec)
    txtRate.text = Format(cmdRate, "##0")
    If DateDiff("s", LastCmdDTS, Now) > CLng(AK_timeout) Then
        If sockServer.State > sckListening Then
            txtStatus(0).text = NowPrefixString & "Disconnect due to Timeout" _
                & vbCrLf _
                & txtStatus(0).text
            DisconnectAK
            AK_RemCntrl.Active = False
            AK_RemCntrl.Run = False
            Write_ELog "AK Comm Link Disconnected due to Timeout"
        End If
        LastCmdDTS = Now
    End If
    If (LocalPagControl.Type = pagMaster) Then
        PaComm_Flag = IIf((sockServer.State = sckConnected), True, False)
    End If
End Sub

Private Function DtsPrefixString() As String
    Dim strDTS, strMS As String
    strMS = Format(AKcmd_Current.CmdRcvdTimer, "##,##0.000")
    strMS = Mid(strMS, (Len(strMS) - 3), 3)
    strDTS = Format(AKcmd_Current.CmdRcvdDTS, "YYYY MMM D  hh:mm:ss") & strMS & "   "
    DtsPrefixString = strDTS
End Function

Private Function NowPrefixString() As String
    Dim strDTS, strMS As String
    strMS = Format(Timer, "##,##0.000")
    strMS = Mid(strMS, (Len(strMS) - 3), 3)
    strDTS = Format(Now(), "YYYY MMM D  hh:mm:ss") & strMS & "   "
    NowPrefixString = strDTS
End Function

Private Sub Cmd_Parse(ByVal cmdIsAK As Boolean, ByVal cmdStr As String)
    Dim paramsStr, curChar, tmpStr, beforeChar, afterChar As String
    Dim iChar, iParam, Idx, max As Integer
    AKcmd_Last = AKcmd_Current
    AKrsp_Last = AKrsp_Current
    ClearCommand
    ClearResponse
    With AKcmd_Current
        ' set CommandReceived
        .CmdRcvdDTS = Now()
        .CmdRcvdTimer = Timer
        ' command is an AK command?
        If Not cmdIsAK Then
            ' Is NOT AK
            tmpStr = cmdStr
            ' remove any non-printable characters
            max = Len(tmpStr)
            For Idx = 1 To max
                iChar = 1 + max - Idx
                curChar = Mid(tmpStr, iChar, 1)
                If ((Asc(curChar) < 32) Or (Asc(curChar) > 126)) Then
                    ' is this the only character?
                    If Len(tmpStr) = 1 Then
                        ' REPLACE non-printable character
                        tmpStr = "nothing"
                    ElseIf Len(tmpStr) = iChar Then
                        ' REMOVE non-printable character at the end of the string
                        tmpStr = Mid(tmpStr, 1, (Len(tmpStr) - 1))
                    ElseIf iChar = 1 Then
                        ' REMOVE non-printable character at the beginning of the string
                        tmpStr = Mid(tmpStr, 2, (Len(tmpStr) - 1))
                    Else
                        ' REMOVE non-printable character from the string
                        beforeChar = Mid(tmpStr, 1, (iChar - 1))
                        afterChar = Mid(tmpStr, (iChar + 1), (Len(tmpStr) - iChar))
                        tmpStr = beforeChar & afterChar
                    End If
                End If
            Next Idx
            If (Len(tmpStr) < 7) Then tmpStr = tmpStr & "-junk--"
            If (Len(tmpStr) > 24) Then tmpStr = Mid(tmpStr, 1, 24)
            .Printable = tmpStr
            .cmdCode = "????"
            .FUnumber = 0
            .NumParams = 0
            paramsStr = ""
            iParam = 0
            tmpStr = ""
        Else
            ' Is AK
            .Printable = Mid(cmdStr, 3, (Len(cmdStr) - 3))
            .cmdCode = Mid(cmdStr, 3, 4)
            If IsNumeric(Mid(cmdStr, 9, 1)) Then .FUnumber = Int(Mid(cmdStr, 9, 1))
            paramsStr = ""
            iParam = 0
            tmpStr = ""
            If (Len(cmdStr) > 10) Then
                paramsStr = Mid(cmdStr, 11, (Len(cmdStr) - 11))
                max = Len(paramsStr)
                For iChar = 1 To max
                    curChar = Mid(paramsStr, iChar, 1)
                    If (curChar = ".") Or (IsNumeric(curChar)) Then
                        tmpStr = tmpStr & curChar
                    Else
                        If IsNumeric(tmpStr) Then
                            iParam = iParam + 1
                            .NumParams = iParam
                            .Param(iParam) = CDbl(tmpStr)
                            tmpStr = ""
                        End If
                    End If
                Next iChar
                If IsNumeric(tmpStr) Then
                    iParam = iParam + 1
                    .NumParams = iParam
                    .Param(iParam) = CDbl(tmpStr)
                    tmpStr = ""
                End If
            End If
            ' inspect the command
            If CommandInterpreted Then
                ' command was interpreted
                .CmdRead = True
                If CommandValid Then
                    ' command is valid (no data errors)
                    .CmdValid = True
                    If CommandAvailable Then
                        ' command is available on this machine
                        .CmdAvailable = True
                        If CommandAccepted Then
                            ' command accepted for processing
                            .CmdAccepted = True
                        End If
                    End If
                End If
            Else
                ' not a known AK command
                .cmdCode = "????"
                .NumParams = 0
            End If
        End If
    End With
End Sub

Private Function CommandDesc(ByVal cmd As String) As String
    Dim str As String
    If Len(cmd) < 10 Then
        str = "is less than 10 characters long"
    ElseIf Len(cmd) > 19 Then
        str = "is more than 19 characters long"
    ElseIf Mid(cmd, 1, 1) <> startChar Then
        str = "does not begin with STX"
    ElseIf InStr(2, cmd, startChar) > 0 Then
        str = "has more than one STX"
    ElseIf Mid(cmd, 2, 1) <> anyChar Then
        str = "has an invalid 2nd character"
    ElseIf Mid(cmd, Len(cmd), 1) <> endChar Then
        str = "does not end with ETX"
    ElseIf InStr(1, cmd, endChar) < Len(cmd) Then
        str = "has more than one ETX"
    Else
        str = "AK"
    End If
    CommandDesc = str
End Function

Private Function CommandInterpreted() As Boolean
    Dim flag As Boolean
    Dim Idx As Integer
    flag = False
    For Idx = 1 To maxCmdIdx
        If AKcmd_Current.cmdCode = AK_Commands(Idx).cmdCode Then
            AKcmd_Current.CmdIndex = Idx
            flag = True
        End If
    Next Idx
    CommandInterpreted = flag
End Function

Private Function CommandValid() As Boolean
    Dim flag As Boolean
    flag = True
    If AKcmd_Current.NumParams < AK_Commands(AKcmd_Current.CmdIndex).MinNumParams Then flag = False
    If AKcmd_Current.NumParams > AK_Commands(AKcmd_Current.CmdIndex).MaxNumParams Then flag = False
    If flag Then
        Select Case AKcmd_Current.cmdCode
            Case "SREQ", "SNRQ"
            Case "ACFG"
            Case "ASTZ"
            Case "ATEM"
            Case Else
        End Select
    End If
    CommandValid = flag
End Function

Private Function CommandAvailable() As Boolean
    Dim flag As Boolean
    flag = True
    If Not AK_Commands(AKcmd_Current.CmdIndex).Available Then flag = False
    CommandAvailable = flag
End Function

Private Function CommandAccepted() As Boolean
    Dim flag As Boolean
    flag = True
    Select Case AKcmd_Current.cmdCode
        Case "SREQ", "SNRQ"
        Case "ACFG"
        Case "ASTZ"
        Case "ATEM"
        Case Else
    End Select
    CommandAccepted = flag
End Function

Private Sub Cmd_Responder()
'
Dim strResponse, strParams As String
Dim Idx, idx1, idx2, vlv As Integer
Dim numZero As Integer
    
    strParams = ""
    numZero = 0
    ' can the command be read ??
    ' (i.e. correct syntax and Command Code is known)
    If Not AKcmd_Current.CmdRead Then
        AKrsp_Current.NumParams = 0
        AKrsp_Current.ParamsType = 2
        AKrsp_Current.ParamStr(1) = ""
    ' are the Parameters valid for the command Code ??
    ElseIf Not AKcmd_Current.CmdValid Then
        AKrsp_Current.NumParams = 1
        AKrsp_Current.ParamsType = 2
        AKrsp_Current.ParamStr(1) = "SE"
    ' does this system respond to this command ??
    ElseIf Not AKcmd_Current.CmdAvailable Then
        AKrsp_Current.NumParams = 1
        AKrsp_Current.ParamsType = 2
        AKrsp_Current.ParamStr(1) = "NA"
    ' can the command be accepted at this time ??
    ElseIf Not AKcmd_Current.CmdAccepted Then
        AKrsp_Current.NumParams = 1
        AKrsp_Current.ParamsType = 2
        If ((LocalPagControl.Type = pagMaster) And (MasterPagData.Status = "SOFF")) Then
            AKrsp_Current.ParamStr(1) = "OF"
        Else
            AKrsp_Current.ParamStr(1) = "BS"
        End If
    Else
        ' respond to the command
        Select Case AKcmd_Current.cmdCode
                
            Case "ACFG"
                ' send all pas configuration values
                AKrsp_Current.ParamsType = 0
                AKrsp_Current.NumParams = 4
                AKrsp_Current.ParamNum(1) = CDbl(SysConfig.Temp_Target)
                AKrsp_Current.ParamNum(2) = CDbl(SysConfig.Moisture_Target)
                AKrsp_Current.ParamNum(3) = CDbl(SysConfig.Tol_Temp)
                AKrsp_Current.ParamNum(4) = CDbl(SysConfig.Tol_Moisture)
                
            Case "ASTZ"
                ' send status
                AKrsp_Current.ParamsType = 2
                AKrsp_Current.NumParams = 3
                AKrsp_Current.ParamStr(1) = IIf((Len(MasterPagData.Status) = 4), MasterPagData.Status, "????")
                AKrsp_Current.ParamStr(2) = IIf(MasterPagData.ReqIn, "1", "0")
                AKrsp_Current.ParamStr(3) = IIf(MasterPagData.RdyOut, "1", "0")
                AKrsp_Current.ParamsType = 2
                
                
            Case "ATEM"
                ' send current pag values
                AKrsp_Current.ParamsType = 0
                AKrsp_Current.NumParams = 4
                AKrsp_Current.ParamNum(1) = CDbl(MasterPagData.Temperature)
                AKrsp_Current.ParamNum(2) = CDbl(MasterPagData.Humidity)
                AKrsp_Current.ParamNum(3) = CDbl(MasterPagData.Moisture)
                
            Case "SNRQ"
                ' requestIn = false
                Select Case LocalPagControl.Type
                    Case pagClient
                        LocalPagControl.ReqIn = False
                    Case pagMaster
                        MasterPagData.ReqIn = False
                    Case Else
                        ' na
                End Select
                AKrsp_Current.NumParams = 0
                
            Case "SREQ"
                ' requestIn = true
                Select Case LocalPagControl.Type
                    Case pagClient
                        LocalPagControl.ReqIn = True
                    Case pagMaster
                        MasterPagData.ReqIn = True
                    Case Else
                        ' na
                End Select
                AKrsp_Current.NumParams = 0
                
            Case Else
                AKrsp_Current.NumParams = 0
        
        End Select
    End If
    ' Response Code is an echo of the Command Code
    AKrsp_Current.RspCode = AKcmd_Current.cmdCode
    ' update & get current error status
    AKrsp_Current.ErrorStatus = ErrorValue_Current
    AKrsp_Current.Printable = AKrsp_Current.RspCode & sepChar & Format(AKrsp_Current.ErrorStatus, "0")
    If (AKrsp_Current.NumParams > 0) Then
        For Idx = 1 To AKrsp_Current.NumParams
            Select Case AKrsp_Current.ParamsType
                Case 0
                    ' the parameter is a floating point
                    strParams = strParams & sepChar & Format(AKrsp_Current.ParamNum(Idx), "####0.000")
                Case 1
                    ' the parameter is an integer
                    strParams = strParams & sepChar & Format(AKrsp_Current.ParamNum(Idx), "####0")
                Case 2
                    ' the parameter is a string
                    strParams = strParams & sepChar & AKrsp_Current.ParamStr(Idx)
            End Select
        Next Idx
        AKrsp_Current.Printable = AKrsp_Current.Printable & strParams
    End If
    ' send the response
    strResponse = startChar & sepChar & AKrsp_Current.Printable & endChar
    sockServer.SendData strResponse
    AKrsp_Current.RspSent = True
    ' log the command ??
    If (LogAkCommands) Then Write_AkLog AKcmd_Current.CmdRcvdDTS, AKcmd_Current.CmdRcvdTimer, AKcmd_Current.Printable, AKrsp_Current.Printable
End Sub

Private Sub ClearCommand()
    Dim Idx As Integer
    With AKcmd_Current
        .Printable = ""
        .FUnumber = 0
        .cmdCode = ""
        .NumParams = 0
        For Idx = 1 To 10
            .Param(Idx) = 0#
        Next Idx
        .CmdType = " "
        .CmdIndex = 0
        .CmdRead = False
        .CmdValid = False
        .CmdAvailable = False
        .CmdAccepted = False
        .CmdRcvdDTS = 0
        .CmdRcvdTimer = 0
    End With
End Sub

Private Sub ClearResponse()
    Dim Idx As Integer
    With AKrsp_Current
        .Printable = ""
        .RspCode = ""
        .ErrorStatus = 0
        .NumParams = 0
        .ParamsType = 0
        For Idx = 1 To 100
            .ParamNum(Idx) = 0#
            .ParamStr(Idx) = ""
        Next Idx
        .RspSent = False
    End With
End Sub

Private Sub InitAKserver()
'
Dim Idx As Integer

    maxCmdIdx = 100
    
    ' initialize AK Command Definition array
    For Idx = 1 To maxCmdIdx
        AK_Commands(Idx).cmdCode = ""
        AK_Commands(Idx).MaxNumParams = 0
        AK_Commands(Idx).MinNumParams = 0
        AK_Commands(Idx).ParamsType = 0
        AK_Commands(Idx).Available = False
    Next Idx
    
    ' AK Command Definitions
    Idx = 0
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = "SREQ"
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = True
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = "SNRQ"
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = True
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = "ACFG"
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = True
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = "ASTZ"
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = True
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = "ATEM"
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = True
    Idx = Idx + 1
    AK_Commands(Idx).cmdCode = ""
    AK_Commands(Idx).MaxNumParams = 0
    AK_Commands(Idx).MinNumParams = 0
    AK_Commands(Idx).ParamsType = 0
    AK_Commands(Idx).Available = False
    
    ' build string of "Known" AK commands
    KnownAKcommands = "_"
    For Idx = 1 To maxCmdIdx
        If AK_Commands(Idx).cmdCode <> "" Then KnownAKcommands = KnownAKcommands & AK_Commands(Idx).cmdCode & "_"
    Next Idx

    ' sockServer state descriptions
    WinsockStateDesc(sckClosed) = "Connection Closed"
    WinsockStateDesc(sckOpen) = "Open"
    WinsockStateDesc(sckListening) = "Listening for Incoming Connections"
    WinsockStateDesc(sckConnectionPending) = "Connection Pending"
    WinsockStateDesc(sckResolvingHost) = "Resolving Remote Client Name"
    WinsockStateDesc(sckHostResolved) = "Remote Client Name Successfully Resolved"
    WinsockStateDesc(sckConnecting) = "Connecting to Remote Client"
    WinsockStateDesc(sckConnected) = "Connected to Remote Client"
    WinsockStateDesc(sckClosing) = "Connection is Closing"
    WinsockStateDesc(sckError) = "Error Occured"

End Sub


