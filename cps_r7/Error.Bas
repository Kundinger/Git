Attribute VB_Name = "Module17"
' error module 17  ''''''''''''''''''' program ERROR.bas '''''''''''''''''''''''''''''
Option Explicit
'
Function ErrorHandler(ErrorIn As Integer) As Integer

' Function Name:    ErrorHandler
' Description:      This routine is used to trap any errors
'                   Each routine has a module identifier and a level
'                   identifier set at the beginning of the routine and
'                   cleared at the end.  The stack may be 100 levels deep.
'
'                   When an error occurs, it is logged, along with the
'                   module and level to help identify its cause later.
'
'                   Some errors allow operation to continue, others require
'                   user input.
'
Dim sMsg As String
Dim iType As Integer
Dim sTitle As String
iType = vbCritical + vbAbortRetryIgnore
If ShortTermErrorCounter < ShortTermErrorMax Then ShortTermErrorCounter = ShortTermErrorCounter + 1
' Write to Event Log
If ElogInitDone Then Write_ELog "Error: " & ErrorIn & _
  ", M" & ErrModule(0) & "-L" & ErrLevel(0) & " " & error$(ErrorIn)
UnreadProgramErrorMessage = True
Select Case ErrorIn
     Case 1 To 6, 9 To 11, 13 To 49, 59, 90, 94             ' Program Related Errors
        sTitle = "Program Error!"
        sMsg = "An unexpected program error has occured." & vbCrLf & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sMsg = sMsg & "Please write down the number and message " & vbCrLf
        sMsg = sMsg & "and notify Analytical Process Systems at " & vbCrLf
        sMsg = sMsg & "1-248-393-0700 to report the error. "
        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case 7                                                  ' Out of Memory
        sTitle = "Out of Memory Error!"
        sMsg = "Not enough memory to complete the operation!" & vbCrLf
        sMsg = sMsg & "Close other programs and try again."
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case 51                                                 ' Internal System Error
        sTitle = "Internal System Error!"
        sMsg = "An unexpected system error has occured." & vbCrLf & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sMsg = sMsg & "Please write down the number and message " & vbCrLf
        sMsg = sMsg & "and notify Analytical Process Systems at " & vbCrLf
        sMsg = sMsg & "1-248-393-0700 to report the error. "
        MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case 52 To 56, 58, 60, 62 To 66, 69, 70, 73 To 76       ' Disk Errors
        sTitle = "System File Error!"
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        ErrorHandler = MsgBox(sMsg, iType, sTitle)                  ' Let user select desired action
    Case 61                                                 ' Disk is full
        sMsg = "Disk is Full!" & vbCrLf & vbCrLf
        sMsg = sMsg & "Change disks or delete " & vbCrLf
        sMsg = sMsg & "unwanted files!" & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sTitle = "Disk Full Error!"
        ErrorHandler = MsgBox(sMsg, iType, sTitle)                  ' Let user select desired action
    Case 57, 68, 71 To 72                                   ' Disk not ready errors
        sMsg = "Disk Not ready or Unavailable!" & vbCrLf & vbCrLf
        sMsg = sMsg & "Check the disk before continuing." & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sTitle = "System Disk Error!"
        ErrorHandler = MsgBox(sMsg, iType, sTitle)               ' Let user select desired action
    Case 67                                                 ' Too many files open
        sMsg = "Not enough file handles to complete the operation!" & vbCrLf
        sMsg = sMsg & "Increase the FILES = number in the config.sys file." & vbCrLf & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sTitle = "File Handles Error!"
        ErrorHandler = MsgBox(sMsg, iType, sTitle)              ' Let user select desired action
    Case 281 To 297                                         ' DDE Errors      No Choices
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sTitle = "System DDE Error!"
        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case 340 To 344                                         ' Control Array Errors      No Choices
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sTitle = "Control Array Error!"
        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case 3021, 3265                                         ' Data base write errors
        sTitle = "Data Base Error Condition on a Write!"
        sMsg = "An unexpected DB error condition has occured." & vbCrLf & vbCrLf
        sMsg = sMsg & "All valves are being closed for safety reasons." & vbCrLf & vbCrLf
        Reset_Valves
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sMsg = sMsg & "Please write down the number and message " & vbCrLf
        sMsg = sMsg & "and notify Analytical Process Systems at " & vbCrLf
        sMsg = sMsg & "1-248-393-0700 to report the error. "
        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
     Case 8000 To 8020                                      ' MsComm Errors
        sTitle = "MsComm Error!"
        sMsg = "An MsComm error has occured for Comm Port #" & CStr(CurCommPort) & "   " & vbCrLf & vbCrLf
        sMsg = sMsg & "Error " & CStr(ErrorIn) & " - "
        Select Case ErrorIn
            Case 8000
                sMsg = sMsg & "Invalid operation on an opened port"
            Case 8001
                sMsg = sMsg & "Timeout value must be greater than zero"
            Case 8002
                sMsg = sMsg & "Invalid port number"
            Case 8003
                sMsg = sMsg & "Property available only at run-time"
            Case 8004
                sMsg = sMsg & "Property is read-only at run-time"
            Case 8005
                sMsg = sMsg & "Port is already open"
            Case 8006
                sMsg = sMsg & "Device identifier is invalid"
            Case 8007
                sMsg = sMsg & "Unsupported baud rate"
            Case 8008
                sMsg = sMsg & "Invalid byte size"
            Case 8009
                sMsg = sMsg & "Error in default parameters"
            Case 8010
                sMsg = sMsg & "Hardware is not available"
            Case 8011
                sMsg = sMsg & "Cannot allocate the queues"
            Case 8012
                sMsg = sMsg & "Device is not open"
            Case 8013
                sMsg = sMsg & "Device is already open"
            Case 8014
                sMsg = sMsg & "Could not enable Comm notification"
            Case 8015
                sMsg = sMsg & "Could not set Comm state"
            Case 8016
                sMsg = sMsg & "Could not set Comm event mask"
            Case 8017
                sMsg = sMsg & "undefined comm error 8017"
            Case 8018
                sMsg = sMsg & "Operation valid only when the port is open"
            Case 8019
                sMsg = sMsg & "Device busy"
            Case 8020
                sMsg = sMsg & "Error reading Comm device"
        End Select
        Write_ELog sMsg
        sMsg = sMsg & vbCrLf & vbCrLf
        If ErrModule(0) <> 4949 Then sMsg = sMsg & "Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf & vbCrLf
        sMsg = sMsg & "Please write down this message " & vbCrLf
        sMsg = sMsg & "and notify Analytical Process Systems at " & vbCrLf
        sMsg = sMsg & "1-248-393-0700 to report the error. "
'        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
    Case Else                                               ' Report any others to APS
        sTitle = "General Error Condition!"
        sMsg = "An unexpected error condition has occured." & vbCrLf & vbCrLf
        sMsg = sMsg & "Error Number: " & CStr(ErrorIn)
        sMsg = sMsg & " Occured at M" & ErrModule(0) & "-L" & ErrLevel(0) & vbCrLf
        sMsg = sMsg & error$(ErrorIn) & vbCrLf & vbCrLf
        sMsg = sMsg & "Please write down the number and message " & vbCrLf
        sMsg = sMsg & "and notify Analytical Process Systems at " & vbCrLf
        sMsg = sMsg & "1-248-393-0700 to report the error. "
        If Not ErrorMsgBypassActive Then MsgBox sMsg, vbInformation, sTitle
        ErrorHandler = vbAbort
End Select

End Function

Sub ChgErrModule(iMod As Integer, iLev As Integer)
'
' Function Name:    ChgErrModule
' Author:           Brunrose         2/05
' Description:      This routine is used with the error handler to
'                   identify errors which might occur during runtime.
'
'                   Each routine has a module identifier and a level
'                   identifier set at the beginning of the routine and
'                   cleared at the end.  The stack may be 100 levels deep.
'
'                   When an error occurs, it is logged, along with the
'                   module and level to help identify its cause later.
'
'                   This routine is called at the middle of a routine
'                   to change the error tracking information for
'                   the current routine.
'
'
' Sets the error tracking information
' iMod is the Module Number
' iLev is the Routine within the Module
'
' The error tracking information allows each routine to be defined by
' module number and subroutine within the module.
'
' Forms are given module numbers starting with 100.
'
' Error messages are logged with the Module and Level number to assist
' in identifying the exact location of the error.
'
' each routine sets the error level on entry,
' optionally changes the error level in the middle of the routine
' and resets the error level on exit.
'
' Up to one hundred levels of error modules are stacked
'
ErrModule(0) = iMod
ErrLevel(0) = iLev

End Sub

Sub SetErrModule(iMod As Integer, iLev As Integer)
'
' Function Name:    SetErrModule
' Author:           Analytical Process Programmer         8/96
' Description:      This routine is used with the error handler to
'                   identify errors which might occur during runtime.
'
'                   Each routine has a module identifier and a level
'                   identifier set at the beginning of the routine and
'                   cleared at the end.  The stack may be 100 levels deep.
'
'                   When an error occurs, it is logged, along with the
'                   module and level to help identify its cause later.
'
'                   This routine is called at the beginning of a routine
'                   to record the error tracking information for
'                   the current routine.
'
'
' Sets the error tracking information
' iMod is the Module Number
' iLev is the Routine within the Module
'
' The error tracking information allows each routine to be defined by
' module number and subroutine within the module.
'
' Forms are given module numbers starting with 100.
'
' Error messages are logged with the Module and Level number to assist
' in identifying the exact location of the error.
'
' each routine sets the error level on entry and resets the error level
' on exit.
'
' Up to one hundred levels of error modules are stacked
'
Dim icnt As Integer

For icnt = 100 To 1 Step -1
  ErrModule(icnt) = ErrModule(icnt - 1)
  ErrLevel(icnt) = ErrLevel(icnt - 1)
Next icnt

ErrModule(0) = iMod
ErrLevel(0) = iLev

End Sub

Sub ResetErrModule()
'
' Function Name:    ResetErrModule
' Author:           Analytical Process Programmer         8/96
' Description:      This routine is used with the error handler to
'                   identify errors which might occur during runtime.
'
'                   Each routine has a module identifier and a level
'                   identifier set at the beginning of the routine and
'                   cleared at the end.  The stack may be 100 levels deep.
'
'                   When an error occurs, it is logged, along with the
'                   module and level to help identify its cause later.
'
'                   THis routine is called at the end of a routine to reset
'                   the error tracking information back one level.
'
Dim icnt As Integer

For icnt = 0 To 99
  ErrModule(icnt) = ErrModule(icnt + 1)
  ErrLevel(icnt) = ErrLevel(icnt + 1)
Next icnt

ErrModule(100) = 0
ErrLevel(100) = 0

End Sub
