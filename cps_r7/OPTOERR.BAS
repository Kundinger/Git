Attribute VB_Name = "module16"
' error module ## none  '''''''''''''program OPTOERR.bas
' Code providec by opto22 brick module
'' OptoErr.BAS
'' This files provides access to the OptoErr DLL.
'' The DLL is used to convert Opto error codes to strings
'' APS 6/1/2000
'' Copyright 1995 Opto 22. All rights reserved.
''---------------------------------------------------------------------------
''
'' Constants for OptoErr.DLL
  Global Const O22_ERROR_MAX_STRING_LENGTH0 = 256 ' 255 chars + 1 null
  Global Const LOWEST_OPTOMUX_ERROR = -8
  Global Const HIGHEST_OPTOMUX_ERROR = -1
'' End of constants for OptoErr.DLL.

'' Function declarations for OptoErr.DLL

'' Given an O22 error code, get the corresponding error string.
Declare Function O22ErrorAsString Lib "OptoErr.DLL" (ByVal ErrorCode%, ByVal DestBuffer$, ByVal MaxBufferSize0&, ActualSize0&) As Integer
' Returns non-zero if OK
' ErrorCode% [in] the error code
' DestBuffer$ [out] Buffer to receive string. May be NULL.
' MaxBufferSize0& [in] max size of buffer including \0.
'ActualSize0& [out] actual size of buffer w/ \0. May be NULL.

'' Given an O22 error code, display the corresponding error string in a pop-up dialog.
Declare Function O22ErrorAsDialog Lib "OptoErr.DLL" (ByVal hwndOwner&, ByVal ErrorCodeArg%, ByVal TextBefore$, ByVal TextAfter$) As Integer
' hwndOwner& [in] owner window handle. May be NULL.
' ErrorCode% [in] the error code
' TextBefore$ [in] Text to be shown before error text. May be NULL.
' TextAfter$ [in] Text to be shown after  error text. May be NULL.

'' End function declarations for OptoErr.DLL

