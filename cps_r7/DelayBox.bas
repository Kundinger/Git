Attribute VB_Name = "Module9"
' error module 9 ''''''''''''''program DELAYBOX.bas ''''''''''''''''''''''
Option Explicit
'
Sub Delay_Box(ByVal Message As String, ByVal tDelay As Integer, ByVal showScreen As Boolean)
'
' Procedure name:   Delay Box
' Author:           Analytical Process Programmer  7/24/96
' Description:      Message Box style box appears, displays message, then
'                   automatically times out and closes.
' tdelay = integer value for number of milliseconds for delay
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9, 1

    If Len(Message) > 40 Then
        frmDelayBox.Width = 5040 * 2
        frmDelayBox.lblMessage.Width = 4695 * 2
    Else
        frmDelayBox.Width = 5040
        frmDelayBox.lblMessage.Width = 4695
    End If
    
     If Not IntroDone Then
         frmDelayBox.Top = frmAbout.Top + 2950
         frmDelayBox.Left = IIf(Len(Message) > 40, frmAbout.Left, frmAbout.Left + 4400)
     ElseIf ShuttingDown Then
         frmDelayBox.Top = frmCheckIt.Top + 2350
         frmDelayBox.Left = IIf(Len(Message) > 40, frmCheckIt.Left, frmCheckIt.Left + 880)
     Else
        Form_Center frmDelayBox
     End If
     
    frmDelayBox.lblMessage = Message
    frmDelayBox.tmrDelayBox.Interval = tDelay   ' in ms
    frmDelayBox.tmrDelayBox.Enabled = True
    If showScreen Then frmDelayBox.Show
    If Len(Message) < 1 Then frmDelayBox.Top = 19999
    DoEvents
    '       Wait for DelayBox to Close
           While frmDelayBox.tmrDelayBox.Enabled = True
    DoEvents
           Wend

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

Sub AckMsg_Box(Message As String)
'
' Procedure name:   AckMsg Box
' Author:           Brunrose 2008
' Description:      Message Box style box appears, displays message, then
'                   waits for user to Acknowledge before it closes.
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 9, 2

If Len(Message) > 40 Then
    frmAckMsg.Width = 5040 * 2
    frmAckMsg.lblMessage.Width = 4695 * 2
Else
    frmAckMsg.Width = 5040
    frmAckMsg.lblMessage.Width = 4695
End If

frmAckMsg.lblMessage = Message
Form_Center frmAckMsg
frmAckMsg.Show

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

