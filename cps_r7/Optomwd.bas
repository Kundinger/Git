Attribute VB_Name = "Module18"
' error module 18 ''''''''''''''''''''' program OPTOMWD.bas '''''''''''''''
' Code provided by OPTO-22
''
'' OptoMwd.Bas
'' This files provides access to the OptoMwd DLL.
'' The DLL is used to communicate with Mistic controllers or Optomux
'' This file [1] defines constants
''           [2] declares functions
'' This header serves as an 'abridged' manual describing all APIs and errors.
''
'----------------------------------------------------------------------------
'     Copyright (C) 1996 Opto 22.
'     All rights reserved.
'----------------------------------------------------------------------------
''
''
''-------------------------------------------------------------------------
'' Constants for OPTOMWD.DLL
''-------------------------------------------------------------------------
Global Const mwdStringLengthMax = &HFF     '' Max length of strings
Global Const mwdProtocolTypeBinary = 1     '' Protocol may be Ascii or binary
Global Const mwdProtocolTypeAscii = 2
Global Const mwdProtocolTypeArcnet = 3     '' Arcnet has its own protocol - neither ascii nor binary
Global Const mwdDataCheckTypeCheckSum = 4  '' dataCheck may be CRC16 or 8 bit checksum
Global Const mwdDataCheckTypeCrc16 = 5
Global Const mwdPortPhysTypeArcnet = 6     '' Arcnet card with SMC compatible chip set
Global Const mwdPortPhysTypeAC37 = 7       '' Opto22's high speed AC37 card
Global Const mwdPortPhysTypeAC39 = 8       '' Not yet supported.
Global Const mwdPortPhysTypeWinApi = 9     '' "regular" RS232 port using Windows APIs.
Global Const ProtocolTypeIOAscii = 10      '' Protocol may now be IO Ascii for Muxes
Global Const ProtocolTypeOptoware = 11     '' New Protocol for OPTOMUX!
Global Const PortPhysTypeAC28 = 12         '' AC28 for Pamux.  Uses section 0 and section 1 of WinRT
Global Const ProtocolTypeIoArcnet = 12     '' Packet for IO mapped arcnet to SNAP
Global Const NoError% = 0
'' End of constants for OPTOMWD.DLL.
''---------------------------------------------------------------------------
'' Function declarations for OPTOMWD.DLL
''---------------------------------------------------------------------------
 Declare Function opto22MwdPortOpenAC37 Lib "c:\windows\system32\OPTOMWD.DLL" (Handle&, ByVal ioPort&, ByVal baud&, ByVal timeOut!, ByVal Retry&, ByVal protocolType&, ByVal datCheckType&) As Integer
 Declare Function opto22MwdPortOpenWinApi Lib "c:\windows\system32\OPTOMWD.DLL" (Handle&, ByVal comPort&, ByVal baud&, ByVal timeOut!, ByVal Retry&, ByVal protocolType&, ByVal datCheckType&) As Integer
 Declare Function opto22MwdPortOpenAC47Io Lib "c:\windows\system32\OPTOMWD.DLL" (Handle&, ByVal ioPort&, ByVal timeOut!, ByVal Retry&, ThisArcnetNodeId As Byte) As Integer
' Used to communicate with optomux.
 Declare Function SendOptoMux Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal iHandle&, ByVal iAddress&, ByVal iCommand&, PositionArray&, ModifierArray&, DataArray&) As Integer
' Used to communicate Directly to bricks
 Declare Function SendMIO Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal iHandle&, ByVal iAddress&, ByVal iCommand&, PositionArray&, SendDataArray&, ReceDataArray&) As Integer
' Extra functions for Visual Basic
' Used to convert binary data from commands such as 'F' or 'TRange.'
 Declare Sub StringAsInt Lib "c:\windows\system32\OPTOMWD.DLL" (NumArg&, ByVal StringArg$)
 Declare Sub StringAsIntMotorola Lib "c:\windows\system32\OPTOMWD.DLL" (NumArg&, ByVal StringArg$)
 Declare Sub StringAsLong Lib "c:\windows\system32\OPTOMWD.DLL" (NumArg&, ByVal StringArg$)
 Declare Sub StringAsSingle Lib "c:\windows\system32\OPTOMWD.DLL" Alias "StringAsFloat" (NumArg!, ByVal StringArg$)
 Declare Sub LongAsString Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal StringArg$, NumArg&)
 Declare Sub SingleAsString Lib "c:\windows\system32\OPTOMWD.DLL" Alias "FloatAsString" (ByVal StringArg$, NumArg!)
''---------------------------------------------------------------------------
'' The following functions are primarily for internal use.
''---------------------------------------------------------------------------
 Declare Function opto22MwdGetVersion2 Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal receive$, ByVal LenMax&) As Integer
 Declare Function opto22MwdSend Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal MisticAdr&, ByVal send$, ByVal sendLength&) As Integer
 Declare Function opto22MwdReceive Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal receive$, ByVal LenMax&, LenActual&) As Integer
 Declare Function opto22MwdReceDataReady Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&) As Integer
 Declare Function opto22MwdReceNoWait Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal receive$, ByVal LenMax&, LenActual&) As Integer
 Declare Function opto22MwdSendRece Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal MisticAdr&, ByVal send$, ByVal sendLength&, ByVal receive$, ByVal LenMax&, LenActual&) As Integer
 Declare Function opto22MwdSendReceNoId Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal MisticAdr&, ByVal send$, ByVal sendLength&, ByVal receive$, ByVal LenMax&, LenActual&) As Integer
 Declare Function opto22MwdPortOpenArcnet Lib "c:\windows\system32\OPTOMWD.DLL" (Handle&, ByVal ioPort&, ByVal memAddr&, ByVal timeOut!, ByVal Retry&) As Integer
 Declare Sub opto22MwdPortClose Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&)
 Declare Sub opto22MwdCounterGet Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, sendAttempt&, sendError&, receAttempt&, receError&)
 Declare Sub opto22MwdCounterClear Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&)
 Declare Function opto22MwdPortLock Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal LockStateBool&) As Integer
 Declare Function opto22MwdPortRetrySet Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal Retry&) As Integer
 Declare Function opto22MwdPortDelaySet Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal Handle&, ByVal timeOut!) As Single
 Declare Function opto22MwdChecksum Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal StrArg$, ByVal LenArg&) As Integer
 Declare Function opto22Crc16 Lib "c:\windows\system32\OPTOMWD.DLL" (ByVal StrArg$, ByVal LenArg&) As Integer
 ' Dial a modem. Temporary back-up for the TAPI approach.
 Declare Function opto22MwdDial Lib "c:\windows\system32\OPTOMWD.DLL" ( _
   ByVal Handle&, _
   ByVal pszAtdtEtc$, _
   ByVal pszResult$, _
   ByVal ResultSize& _
 ) As Integer
 ' Handle&, Port Handle.
 ' pszAtdtEtc$, Dial Command
 ' pszResult$, Modem result string, null terminated.
 ' ResultSize&, prevents string overruns.
 ' Returns Opto error.
'' End function declarations for OPTOMWD.DLL.

