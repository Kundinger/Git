Attribute VB_Name = "Module5"
' error Module 5 ''''''''''''''' program  HOTKEY.bas
Option Explicit
'
Sub HotKeyCheck(KeyCode As Integer, Shift As Integer)
'
' Function Name:    HotKeyCheck
' Author:           Analytical Process Programmer         8/96
' Description:      This routine is used to provide Special Function
'                   Keys.
'
'                   Insert a call to HotKeyCheck on a Form's KeyDownEvent
'                   to have the key press intercepted and checked
'                   to see if a HotKey event should be called.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 5, 1
Dim ShiftDown, AltDown, CtrlDown, alldown As Boolean
Dim sString As String
Dim batcmd As String

Const vbShiftMask = 1
Const vbCtrlMask = 2
Const vbAltMask = 4
Const vbF1Mask = 112

    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    alldown = ShiftDown And AltDown And CtrlDown
    
    ' Logon
    If KeyCode = vbKeyP And AltDown And Not CtrlDown And Not ShiftDown Then
        frmPassword.Show
    End If
    
    ' Alt L - Logout
    If KeyCode = vbKeyL And AltDown And Not CtrlDown And Not ShiftDown Then
        Logout
    End If
    
    ' AltCntrlShift A - Edit UserName/Password/AccessCodes
    If KeyCode = vbKeyA And alldown Then
        If CheckPass("5", True) Then frmPassEdit.Show
    End If
    
    ' AltCntrlShift X - Exit
    If KeyCode = vbKeyX And alldown Then
        sString = "You are about to Exit the Program!" & vbCrLf & vbCrLf _
            & "Exiting will stop all currently running tests." & vbCrLf & vbCrLf _
            & "Please confirm Exit; or Cancel"
        frmCheckIt.CheckIt 1, sString
    End If
    
    ' Cntrl F1 - Help
    If KeyCode = vbF1Mask Then
        If CtrlDown And Not AltDown And Not ShiftDown Then
            batcmd = FILEPATH_manuals & "showDoc.bat  cps_r7.pdf"
            Shell batcmd    ' Run bat to open documentation
        Else
            Delay_Box "Use CTRL and F1 for Help Documentation", MSGDELAY, msgSHOW
        End If
    End If
    
    ' Alt A - About
    If KeyCode = vbKeyA And AltDown And Not CtrlDown And Not ShiftDown Then
        If CheckPass("D", True) Then
            frmAbout.Show
        End If
    End If
    
    ' Alt P Shift - System Definition Screen
    If KeyCode = vbKeyP And AltDown And ShiftDown And Not CtrlDown Then
    '    If CheckPass("H", True) Then
            frmSysDefMain.Show
    '    End If
    End If
    
    ' Cntrl P Shift - DataWatcher
    If KeyCode = vbKeyP And CtrlDown And ShiftDown And Not AltDown Then
        If CheckPass("F", True) Then
            frmDataWatcher.Left = frmMainMenu.Left
            frmDataWatcher.Top = frmMainMenu.Top
            frmDataWatcher.Show
        End If
    End If
    
    ' CtrlAlt B - Low Butane
    If KeyCode = vbKeyB And AltDown And CtrlDown And Not ShiftDown Then
        If Not IoComOn And USINGSIMULATION Then ButaneSupply.CurrentOnHand = 1
    End If
    
    ' CtrlAlt T - Testing Screen
    If KeyCode = vbKeyT And AltDown And CtrlDown And Not ShiftDown Then
        frmCourses.Show
    End If
    
    ' Alt F - First Screen
    If KeyCode = vbKeyF And AltDown And Not ShiftDown And Not CtrlDown Then
        frmFirstAid.Show
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
