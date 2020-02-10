Attribute VB_Name = "Module7"
' error module 7 '''''''''''''' program PASSWORD.bas ''''''''''''''''''''''''''
Option Explicit
'
Private daodb36 As DAO.Database
Private rS As DAO.Recordset
Dim sPath As String

Function CheckPass(level As String, showmsg As Boolean) As Boolean
'
' Routine Name:     CheckPass
' Author:           Analytical Process Programmer  APS
' Description:      Checks the passed access key against the current user
'                   access list to determine if current user has access
'                   for desired operation.
'
'                   Routine returns a true value if the key is contained
'                   in the list, false otherwise.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7, 1

If InStr(CurrentUser.Access, level) <> 0 Then
  CheckPass = True
Else
  CheckPass = False
  If showmsg Then Delay_Box "Insufficient Access!", MSGDELAY, msgSHOW
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
Sub Init_Password()
'
' Function Name:    Init_Password
' Author:           Analytical Process Programmer         8/96
' Description:      This routine initializes the Password Access Variables
'
'   ***********
'   Access List
'   ***********
'   User    CPS Admin   APS MASTER
'   ****    *** *****   *** ******
'   A       A   A       A   A = Edit Station Canister/Recipe
'           B   B       B   B = View/Load/Save Configuration Data
'               C       C   C = Controller Config/Tune
'           D   D       D   D = Documentation
'           E   E       E   E = Job DB / Report Copy
'           F   F       F   F = Report Print, Generate, Review
'           G   G       G   G = Exit Program
'                       H   H = Access System Definition
'                       I   I = Access Simulation Controls
'   J       J   J       J   J = Access Password Screen
'   K       K   K       K   K = View Station Detail Screen
'           L   L       L   L = View File Log
'           M   M       M   M = View Job List
'   N       N   N       N   N = View Master Recipes & Canister Recipes
'           O   O       O   O = Load/Save Master Recipes
'           P   P       P   P = Load/Save Master Canister Recipes
'   Q       Q   Q       Q   Q = Change Station Detail Text
'   R       R   R       R   R = Start/Stop Station
'               S       S   S = Data Log Delete/Clear
'               T       T   T = Job Log Delete/Clear
'   U       U   U       U   U = Thermocouples, etc. On / Off on Station Detail
'               V       V   V = Allow Manual File Maintenance
'           W   W       W   W = Printer Font Select Test Screen Button
'           X   X       X   X = Access MFC Calibration Screen
'               Y       Y   Y = Configure advanced config parameters (file maint, AutoLogon, etc.)
'           Z   Z       Z   Z = Access Event Log
'                       0   0 = Change Debug Flags, IoComOn, etc. on SysDefScreen
'               1       1   1 = Change File Name / Job# in config
'           2   2       2   2 = Access I/O Monitor screens
'           3   3       3   3 = Access Scale Monitor Screen
'           4   4       4   4 = unused (reserved for user cps)
'               5       5   5 = Access Password Maintenance Screen
'               6       6   6 = unused (reserved for user admin)
'               7       7   7 = user admin
'                       8   8 = user aps
'                       9   9 = Access System Timer Values

'
' Initialize Password Access
Dim dbDbase As Database
Dim rsTable As Recordset
Dim rsCriterion As String
Dim sPath, sUserName As String
Dim LtrPos As Long
Dim TrimXs As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7, 2

ApsUser.USER = "APS"
ApsUser.PWord = "APS"
ApsUser.Access = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
MasterUser.USER = "BRUNROSE"
MasterUser.PWord = "BRUNROSE"
MasterUser.Access = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
DefaultUser.USER = "DEFAULT"
DefaultUser.PWord = "PASSWORD"
DefaultUser.Access = "JKN"

Select Case AutoLogon
    
    Case autologonOFF
        CurrentUser = DefaultUser
        
    Case autologonON
        Select Case SysConfig.AutoLogon
            Case autologonOFF
                CurrentUser = DefaultUser
            Case Else
                sUserName = SysConfig.AutoLogonUser
                sPath = FILEPATH_sysdbf & DATAUSER
                Set daodb36 = DBEngine.OpenDatabase(sPath)
                Set rS = daodb36.OpenRecordset("password")
                Set frmPassEdit.Data1.Recordset = rS
                
                ' Check password list
                rsCriterion = "SELECT * FROM PASSWORD WHERE " & _
                             "([UserName] = '" & sUserName & "')"
                Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAUSER)
                Set rsTable = dbDbase.OpenRecordset(rsCriterion, dbOpenDynaset)
                If Not rsTable.BOF Then  ' See if valid user exists
                    rsTable.MoveFirst
                    CurrentUser.USER = rsTable("UserName")
                    CurrentUser.PWord = rsTable("PassCode")
                    CurrentUser.Access = rsTable("Access")
                    ' ensure only APS has access to Sysdef
                    If UCase(CurrentUser.USER) <> "APS" Then
                        If InStr(1, CurrentUser.Access, "H", vbTextCompare) > 0 Then
                            LtrPos = InStr(1, CurrentUser.Access, "H", vbTextCompare)
                            TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                            TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                            CurrentUser.Access = TrimXs
                        End If
                        If InStr(1, CurrentUser.Access, "I", vbTextCompare) > 0 Then
                            LtrPos = InStr(1, CurrentUser.Access, "I", vbTextCompare)
                            TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                            TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                            CurrentUser.Access = TrimXs
                        End If
                        If InStr(1, CurrentUser.Access, "0", vbTextCompare) > 0 Then
                             LtrPos = InStr(1, CurrentUser.Access, "0", vbTextCompare)
                            TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                            TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                            CurrentUser.Access = TrimXs
                        End If
                    End If
                    frmAbout.UpdateMsg "Logged on as " & UCase(CurrentUser.USER) & vbCrLf
                Else
                    CurrentUser = DefaultUser
                    frmAbout.UpdateMsg "Failed to find Auto Logon username = " & UCase(sUserName) & vbCrLf
                End If
                rsTable.Close
                dbDbase.Close
        End Select

    Case autologonAPS
        CurrentUser = ApsUser
        frmAbout.UpdateMsg "Logged On as APS" & vbCrLf
        
    Case Else
        AutoLogon = autologonOFF
        CurrentUser = DefaultUser
        
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

Sub Logout()
'
' Routine Name:   Logout
' Author:         Analytical Process Programmer APS
' Description:    Logs current user off of system.
'                 Writes message to the message log.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7, 3

' Write to Event Log,
CurrentUser = DefaultUser
CurrentUser.Access = "BJ"
Write_ELog "User: " & CurrentUser.USER & " Logged Out."
frmPassword.Show
DoEvents
frmPassword.lblMessage.ForeColor = Message_ForeColor
frmPassword.lblMessage.Caption = "User Logged Out!"

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

Function SetPass(sUser As String, sCode As String) As Boolean
' Function Name:    SetPass
' Author:           Analytical Process Programmer         8/96
' Description:      This routine checks the entered password and user id
'                   against the master password and against the password
'                   database.
'
'                   If the user id and password exist and are correct, the
'                   routine sets the current user to the new user info and
'                   returns a true value.
'
'                   If the information does not match, an appropriate
'                   messsage is displayed and the routine returns a false.
'
Dim dbDbase As Database
Dim rsTable As Recordset
Dim rsCriterion As String
Dim sPath As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 7, 4

    ' result is False unless a match is found
    SetPass = False
    
    ' See if pre-defined user
    If UCase(sUser) = MasterUser.USER And UCase(sCode) = MasterUser.PWord Then
        ' Master user
        CurrentUser = MasterUser
        SetPass = True
    ElseIf UCase(sUser) = ApsUser.USER And UCase(sCode) = ApsUser.PWord Then
        ' Aps user
          CurrentUser = ApsUser
          SetPass = True
    Else
    ' Check password list
        sPath = FILEPATH_sysdbf & DATAUSER
        Set daodb36 = DBEngine.OpenDatabase(sPath)
        Set rS = daodb36.OpenRecordset("password")
        Set frmPassEdit.Data1.Recordset = rS
        
        ' does that user exist ?
        rsCriterion = "SELECT * FROM PASSWORD WHERE " & _
                      "([UserName] = '" & sUser & "') AND " & _
                      "([PassCode] = '" & sCode & "')"
        Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAUSER)
        Set rsTable = dbDbase.OpenRecordset(rsCriterion, dbOpenDynaset)
        If Not rsTable.BOF Then
            rsTable.MoveFirst
            CurrentUser.USER = sUser
            CurrentUser.PWord = sCode
            CurrentUser.Access = rsTable("Access")
            ' ensure only APS & Master have access to Sysdef
            Dim LtrPos As Long
            Dim TrimXs As String
            If InStr(1, CurrentUser.Access, "H", vbTextCompare) > 0 Then
                LtrPos = InStr(1, CurrentUser.Access, "H", vbTextCompare)
                TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                CurrentUser.Access = TrimXs
            End If
            If InStr(1, CurrentUser.Access, "I", vbTextCompare) > 0 Then
                LtrPos = InStr(1, CurrentUser.Access, "I", vbTextCompare)
                TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                CurrentUser.Access = TrimXs
            End If
            If InStr(1, CurrentUser.Access, "0", vbTextCompare) > 0 Then
                 LtrPos = InStr(1, CurrentUser.Access, "0", vbTextCompare)
                TrimXs = Mid(CurrentUser.Access, 1, LtrPos - 1)
                TrimXs = TrimXs & Mid(CurrentUser.Access, LtrPos + 1, (Len(CurrentUser.Access) - Len(TrimXs) - 1))
                CurrentUser.Access = TrimXs
            End If
            SetPass = True
        End If
        rsTable.Close
        dbDbase.Close
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
