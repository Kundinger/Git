Attribute VB_Name = "Module14"
'error module 14 ''''''''''''''program UTILS.bas '''''''''''''''''''''''
Option Explicit
'
Public Sub PurgeWizardUpdater(ByVal iTime As Single, ByVal iFlow As Single)
    frmRecipe.Show
    frmRecipe.UpdatePurgeFlowTime iTime, iFlow
'    frmRecipe.Refresh
End Sub

Sub DelayBySeconds(ByVal secs As Double)
Dim TmrEnd, DeltaSecs As Double
    TmrEnd = Timer + secs
    If TmrEnd > 86400 Then Exit Sub
    DoEvents
    Do
        DoEvents
        DeltaSecs = Abs(Timer - TmrEnd)
    Loop Until DeltaSecs > secs
    DoEvents
End Sub

Function Decimal2Binary(ByVal iDecimal As Integer) As String
' Decimal Number to Binary string:
Dim idx As Integer
Dim DecNum As Integer
Dim pow2 As Single
Dim BinStr As String
    DecNum = iDecimal
    BinStr = ""
    For idx = 14 To 0 Step -1           ' Positive numbers only
        pow2 = 2 ^ idx
        If DecNum >= pow2 Then
            DecNum = DecNum - pow2
            BinStr = BinStr + "1"
        Else
            BinStr = BinStr + "0"
        End If
    Next idx
    Decimal2Binary = BinStr
End Function


Function NameOfMonth(ByVal iMonth As Integer) As String
'
' Module Name:  NameOfMonth
' Author:       Brunrose
' Description:  This routine converts an integer between 1 and 12 into
'               the corresponding Month Name string.
'
    Select Case iMonth
        Case 1
            NameOfMonth = "January"
        Case 2
            NameOfMonth = "February"
        Case 3
            NameOfMonth = "March"
        Case 4
            NameOfMonth = "April"
        Case 5
            NameOfMonth = "May"
        Case 6
            NameOfMonth = "June"
        Case 7
            NameOfMonth = "July"
        Case 8
            NameOfMonth = "August"
        Case 9
            NameOfMonth = "September"
        Case 10
            NameOfMonth = "October"
        Case 11
            NameOfMonth = "November"
        Case 12
            NameOfMonth = "December"
        Case Else
            NameOfMonth = "unknown"
    End Select
End Function

Function GramsPerHourToSlpm(ByVal sGrams As Single, ByVal sGramsPerLiter As Single) As Single
'
' Module Name:  GramsPerHourToSlpm
' Author:       Analytical Process Programmer 9/96
' Description:  This routine converts a value provided in Grams Butane
'               per hour to Standard Liters per Minute of Butane FLow.
'
'GramsPerHourToSlpm = sGrams / 144.38
GramsPerHourToSlpm = sGrams / (sGramsPerLiter * 60#)

End Function

Function LineVolume(sID As Single, sLength As Single) As Single
'
' Module Name:  LineVolume
' Author:       Analytical Process Programmer 10/2007
' Description:  This routine calculates Line Volume from ID and Length.
'               The result is always in Liters.
'               The Engr. Units of the input values varies with SI/English.
'
Dim a, r, v As Single
        
    ' vol    =    pie   *    (r squared)   *   length
    If USINGLVol_SI Then
        '
        ' USING SI UNITS
        '
        ' ID in mm;  Length in Meters
        '
        ' 1 millimeters = 0.001 meters
        ' 1 cubic meter = 1000 Liters
        '
        r = (sID / 2) * 0.001       ' in meters
        a = (22 / 7) * (r * r)      ' in square meters
        v = (a * (sLength))         ' in cubic meters
        v = v * 1000                ' in Liters
    ElseIf USINGLVol_Engl Then
        '
        ' USING ENGLISH UNITS
        '
        ' ID in inches;  Length in Feet
        '
        ' 1 cubic inches = 0.016387064 liters
        '
        r = (sID / 2)               ' in inches
        a = (22 / 7) * (r * r)      ' in square inches
        v = (a * (sLength * 12))    ' in cubic inches
        v = v * 0.016387064         ' in Liters
    Else
        '
        ' UNKNOWN UNITS
        '
        v = 0#                      ' in Liters
    End If

    LineVolume = v

End Function

Function RHtoGrains(ByVal sBaro As Single, ByVal sTemp As Single, ByVal sRH As Single) As Single
  
' Module Name:  RHtoGrains
' Author:       Analytical Process Programmer 1/97
' Description:  This routine uses the relative humidity, temperature
'               and barometric pressure values to generate a value of
'               PAMoisture in Grains per pound.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 14, 2380
Dim val1 As Single
Dim val2 As Single
Dim valTemp As Single

    If USINGC Then
       valTemp = sTemp
    ElseIf USINGF Then
       valTemp = ((sTemp - 32) * 5 / 9)  ' convert to deg C
    End If
    val1 = (18.678 - (valTemp / 234.5)) * valTemp / (valTemp + 257.14)
    val2 = 6.1366 * Exp(val1) * (sRH / 100)
    ' prevent divide by zero
    If sBaro = val2 Then sBaro = sBaro + 1
    RHtoGrains = 0.007 * 622000 * val2 / (sBaro - val2)
  
ResetErrModule
Exit Function
 
localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Function Set_StrLen(ByVal InString As String, sLength As Integer, _
  leading As Boolean) As String
'
' Module Name:  Set_StrLen
' Author:       Analytical Process Programmer 9/96
' Description:  This routine adjusts the input string to the length
'               specified.
'
' Leading will place spaces in front of the passed string when true
' and will place spaces behind the passed string when false.
'
' if the input string is longer than length, the returned string is
' chopped to the length passed.
'
SetErrModule 14, 8
If UseLocalErrorHandler Then On Error GoTo localhandler

Dim OutString As String
If Len(InString) < sLength Then
  If leading Then
    Set_StrLen = Space(sLength - Len(InString)) & InString
  Else
   Set_StrLen = InString & Space(sLength - Len(InString))
  End If
Else
  Set_StrLen = Left(InString, sLength)
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

Function Flood_Set(ByVal hilim As Integer, ByVal LoLim As Integer, ByVal MeasVal As Single) As Integer
'
' Module Name:  Flood_Set
' Author:       Analytical Process Programmer 9/96
' Description:  This routine returns a value between 0 and 100 representing
'               the appropriate panel flood value for the measured value
'               supplied, based upon the upper and lower limits.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 14, 3

Dim sresult As Single
If hilim - LoLim <> 0 Then
  sresult = 100 * (MeasVal - LoLim) / (hilim - LoLim)
  Flood_Set = IIf(sresult < 0, 0, IIf(sresult > 100, 100, CInt(sresult)))
Else
  Flood_Set = 100
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

Sub MyFocus(tbox As TextBox)
'
' Module Name:  MyFocus
' Author:       Analytical Process Programmer 9/96
' Description:  This routine sets the focus of a form to the specified
'               text box and highlights the contents.
'
SetErrModule 14, 6

tbox.SelStart = 0
tbox.SelLength = Len(tbox.text)

ResetErrModule
End Sub

Function FloodCheck(sVal As Single) As Single
' Common Utilities Module
' Module Name:  FloodCheck
' Author:       Analytical Process Programmer 9/96
' Description:  This routine checks to ensure the passed value is between
'               0 and 100.
'
SetErrModule 14, 4

If sVal < 0 Then
  FloodCheck = 0
Else
  If sVal > 100 Then
    FloodCheck = 100
  Else
    FloodCheck = sVal
  End If
End If

ResetErrModule
End Function

Function DegCtoF(ctemp As Single) As Single
'
' Module Name:  DegCtoF
' Author:       Analytical Process Programmer 9/96
' Description:  This converts a value in degrees celsius to a value in
'               degrees fahrenheit.
'
SetErrModule 14, 1
DegCtoF = (ctemp * 9 / 5) + 32

ResetErrModule

End Function

Function DegFtoC(fTemp As Single) As Single
'
' Module Name:  DegFtoC
' Author:       Analytical Process Programmer 9/96
' Description:  This converts a value in degrees fahrenheit to a value in
'               degrees celsius.
'
SetErrModule 14, 2
DegFtoC = (fTemp - 32) * (5 / 9)

ResetErrModule
End Function

Sub Form_Center(cForm As Form)
'
' Module Name:  Form_Center
' Author:       Analytical Process Programmer 9/96
' Description:  This routine centers the current form
'
SetErrModule 14, 5

cForm.Top = (Screen.Height - cForm.Height) / 2
cForm.Left = (Screen.Width - cForm.Width) / 2

ResetErrModule
End Sub

Function Range_Check(hiVal As Single, loVal As Single, ByRef tbox As TextBox, _
  sMessage As String) As Boolean
'
' Module Name:  Range_check
' Author:       Analytical Process Programmer 9/96
' Description:  This routine checks the value of the text box supplied and
'               compares to the high and low limits.  If the values is
'               within or equal to the limits, the routine returns a value
'               of TRUE, otherwise a FALSE value.
'
'               When the return value is false, a message supplying the
'               allowable limits is displayed, and the contents of the
'               text box are highlighted.
'
SetErrModule 14, 7
If UseLocalErrorHandler Then On Error GoTo localhandler

    If tbox.text = "" Then tbox.text = "0"
    If CStr(tbox) < loVal Or CStr(tbox) > hiVal Then
      Range_Check = False
      Delay_Box sMessage & " Range Error!" & vbCrLf & "Allowable Range = " _
        & loVal & " - " & hiVal & ".", MSGDELAY, msgSHOW
      MyFocus tbox
    Else
      Range_Check = True
    End If

ResetErrModule
Exit Function

localhandler:
Range_Check = False
  Delay_Box sMessage & " Range Error!" & vbCrLf & "Allowable Range = " _
    & loVal & " - " & hiVal & ".", MSGDELAY, msgSHOW
MyFocus tbox
End Function

Function SlpmToGramsPerHour(ByVal sSlpm As Single, ByVal sGramsPerLiter As Single) As Single
'
' Module Name:  SlpmToGramsPerHour
' Author:       Analytical Process Programmer 9/96
' Description:  This routine converts a value provided in
'               SLPM of Butane flow
'               to Butane Flow in Grams Butane per Hour
'
'SlpmToGramsPerHour = sSlpm * 144.38
SlpmToGramsPerHour = sSlpm * (sGramsPerLiter * 60#)

End Function

Function Verify_Float(entry As String, KeyAscii As Integer) As Integer
'
' Module Name:  Verify_Float
' Author:       Analytical Process Programmer 9/96
' Description:  This routine is used in text box keypress events to allow
'               only floating point valid characters.
'
SetErrModule 14, 9
If UseLocalErrorHandler Then On Error GoTo localhandler

    Select Case KeyAscii
      Case 8, 9, 10, 13, 48 To 57
        Verify_Float = KeyAscii
      Case 46 ' Decimal Point
        If InStr(entry, ".") Then Verify_Float = 0 Else Verify_Float = KeyAscii
      Case Else
        Verify_Float = 0
    End Select

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

Function Verify_Int(entry As String, KeyAscii As Integer) As Integer
'
' Module Name:  Verify_Int
' Author:       Analytical Process Programmer 9/96
' Description:  This routine is used in text box keypress events to allow
'               only integer type valid characters.
'
SetErrModule 14, 10
If UseLocalErrorHandler Then On Error GoTo localhandler
    Select Case KeyAscii
      Case 8, 9, 10, 13, 48 To 57
        Verify_Int = KeyAscii
      Case Else
        Verify_Int = 0
    End Select

ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Function

Public Function IsEven(MyInt As Integer)
    ' Returns True if MyInt is an even number
    
    Dim a As Integer
    Dim b As Integer
    
    a = MyInt / 2
    b = a * 2
    
    If b = MyInt Then
        IsEven = True
    Else
        IsEven = False
    End If
End Function

Public Function IsOdd(MyInt As Integer)
    ' Returns True if MyInt is an odd number
    
    Dim a As Integer
    Dim b As Integer
    
    a = MyInt / 2
    b = a * 2
    
    If b = MyInt Then
        IsOdd = False
    Else
        IsOdd = True
    End If
End Function

Public Function Thermocouple(engineeringunits As Variant, tctype As Single) As Variant
'
' Module Name:  Thermocouple
' Author:       Analytical Process Programmer 11/02
' Description:  This routine converts engineering units to temperature C/F depending
'               on Thermocouple type case statement and Metric mode
'
' Thermocouple constants
' TCTypeB = 0            ' Range = +42C to +1820
' TCTypeJ = 1            ' Range = -270C to +1200
' TCTypeK = 2            ' Range = -270C to +1372
' TCTypeR = 3            ' Range = -50C to +1768
' TCTypeS = 4            ' Range = -50C to +1768
' TCTypeT = 5            ' Range = -270C to +400

    Select Case tctype
      Case TCTypeB       '
        
      Case TCTypeJ       '
       ' ADJUST raw inputs type J
          Thermocouple = ((((engineeringunits / 65536) * 9) / 5) + 32) ' offset by requester + TC_Offset
      Case TCTypeK       '
      
      Case TCTypeR       '
        
      Case TCTypeS       '
        
      Case TCTypeT       '
        
    End Select
End Function

Public Function Bits_Pack(Optional B00 As Boolean, _
                        Optional B01 As Boolean, _
                        Optional B02 As Boolean, _
                        Optional B03 As Boolean, _
                        Optional B04 As Boolean, _
                        Optional B05 As Boolean, _
                        Optional B06 As Boolean, _
                        Optional B07 As Boolean, _
                        Optional B08 As Boolean, _
                        Optional B09 As Boolean, _
                        Optional B10 As Boolean, _
                        Optional B11 As Boolean, _
                        Optional B12 As Boolean, _
                        Optional B13 As Boolean, _
                        Optional B14 As Boolean) As Integer
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 14, 1621
Dim tmp, idx As Integer

    tmp = 0
    For idx = 0 To 14
        Select Case idx
            Case 0
                If B00 Then tmp = tmp + 1
            Case 1
                If B01 Then tmp = tmp + 2
            Case 2
                If B02 Then tmp = tmp + 4
            Case 3
                If B03 Then tmp = tmp + 8
            Case 4
                If B04 Then tmp = tmp + 16
            Case 5
                If B05 Then tmp = tmp + 32
            Case 6
                If B06 Then tmp = tmp + 64
            Case 7
                If B07 Then tmp = tmp + 128
            Case 8
                If B08 Then tmp = tmp + 256
            Case 9
                If B09 Then tmp = tmp + 512
            Case 10
                If B10 Then tmp = tmp + 1024
            Case 11
                If B11 Then tmp = tmp + 2048
            Case 12
                If B12 Then tmp = tmp + 4096
            Case 13
                If B13 Then tmp = tmp + 8192
            Case 14
                If B14 Then tmp = tmp + 16384
            Case Else
                ' nothing to do
        End Select
    Next idx
    Bits_Pack = tmp

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

Public Function Bits_UnPack(ByVal WordOfBits As Integer) As SixteenBits
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 14, 1216
Dim tmp, idx, idx2 As Integer
Dim outBits As SixteenBits

    tmp = WordOfBits
    For idx2 = 0 To 14
        idx = 14 - idx2
        Select Case idx
            Case 0
                outBits.B00 = IIf(tmp >= 1, True, False)
                If tmp >= 1 Then tmp = tmp - 1
            Case 1
                outBits.B01 = IIf(tmp >= 2, True, False)
                If tmp >= 2 Then tmp = tmp - 2
            Case 2
                outBits.B02 = IIf(tmp >= 4, True, False)
                If tmp >= 4 Then tmp = tmp - 4
            Case 3
                outBits.B03 = IIf(tmp >= 8, True, False)
                If tmp >= 8 Then tmp = tmp - 8
            Case 4
                outBits.B04 = IIf(tmp >= 16, True, False)
                If tmp >= 16 Then tmp = tmp - 16
            Case 5
                outBits.B05 = IIf(tmp >= 32, True, False)
                If tmp >= 32 Then tmp = tmp - 32
            Case 6
                outBits.B06 = IIf(tmp >= 64, True, False)
                If tmp >= 64 Then tmp = tmp - 64
            Case 7
                outBits.B07 = IIf(tmp >= 128, True, False)
                If tmp >= 128 Then tmp = tmp - 128
            Case 8
                outBits.B08 = IIf(tmp >= 256, True, False)
                If tmp >= 256 Then tmp = tmp - 256
            Case 9
                outBits.B09 = IIf(tmp >= 512, True, False)
                If tmp >= 512 Then tmp = tmp - 512
            Case 10
                outBits.B10 = IIf(tmp >= 1024, True, False)
                If tmp >= 1024 Then tmp = tmp - 1024
            Case 11
                outBits.B11 = IIf(tmp >= 2048, True, False)
                If tmp >= 2048 Then tmp = tmp - 2048
            Case 12
                outBits.B12 = IIf(tmp >= 4096, True, False)
                If tmp >= 4096 Then tmp = tmp - 4096
            Case 13
                outBits.B13 = IIf(tmp >= 8192, True, False)
                If tmp >= 8192 Then tmp = tmp - 8192
            Case 14
                outBits.B14 = IIf(tmp >= 16384, True, False)
                If tmp >= 16384 Then tmp = tmp - 16384
        End Select
    Next idx2
    Bits_UnPack = outBits

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

Public Function RedFromRGB(ByVal rgb As Long) As Integer
' The ampersand after &HFF coerces the number as a
' long, preventing Visual Basic from evaluating the
' number as a negative value. The logical And is
' used to return bit values.
    RedFromRGB = &HFF& And rgb
End Function

Public Function GreenFromRGB(ByVal rgb As Long) As Integer
' The result of the And operation is divided by
' 256, to return the value of the middle bytes.
' Note the use of the Integer divisor.
    GreenFromRGB = (&HFF00& And rgb) \ 256
End Function

Public Function BlueFromRGB(ByVal rgb As Long) As Integer
' This function works like the GreenFromRGB above,
' except you don't need the ampersand. The
' number is already a long. The result divided by
' 65536 to obtain the highest bytes.
    BlueFromRGB = (&HFF0000 And rgb) \ 65536
End Function

Public Function RGBFromRedGreenBlue(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long) As Long
' This function works like the GreenFromRGB above,
' except you don't need the ampersand. The
' number is already a long. The result divided by
' 65536 to obtain the highest bytes.
Dim Hred, Hgreen, Hblue As Long
    Hred = &HFF& And Red
    Hgreen = (&HFF& And Green) * 256
    Hblue = (&HFF And Blue) * 65536
    RGBFromRedGreenBlue = Hred + Hgreen + Hblue
End Function

Public Function MinutesFromNow(ByVal nummin As Single) As Date
' Routine Name: MinutesFromNow
' Created by:   Brunrose
' Function:
' This routine is TimeSerial with decimal minutes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim iMin, iSec As Integer

SetErrModule 14, 9090
If UseLocalErrorHandler Then On Error GoTo localhandler

    iMin = CInt(nummin)
    If iMin > nummin Then iMin = iMin - 1
    iSec = IIf(nummin > iMin, CInt((60# * (nummin - CSng(iMin)))), 0)
    
    MinutesFromNow = TimeSerial(0, iMin, iSec)

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

Public Function DiffTimer(ByVal inTimer As Double) As Double
' Routine Name: DiffTimer
' Created by:   Brunrose
' Function:
' This routine returns the difference in sec.millisec between the
' current Timer value and the specified value.
' If (DiffTimer > 0) then specified time is in the future.
' If (DiffTimer < 0) then specified time is in the past.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim iDiff As Double
Dim iTimr As Double
SetErrModule 14, 1192
If UseLocalErrorHandler Then On Error GoTo localhandler

    ' current value of Timer
    iTimr = Timer
    ' valid inTimer ??
    If inTimer = Empty Then
        iDiff = 0
    ElseIf Not IsNumeric(inTimer) Then
        iDiff = 0
    ElseIf inTimer < 0 Then
        iDiff = 0
    ElseIf inTimer > 86399.999 Then
        iDiff = 0
    Else
        ' calculate difference in seconds.ms
        Select Case inTimer
            Case Is < iTimr
                ' in past; diff > 0
                iDiff = iTimr - inTimer
            Case Is = iTimr
                ' now; diff=0
                iDiff = 0
            Case Is > iTimr
                ' in future; diff < 0
                iDiff = Timer - inTimer
        End Select
    End If
    DiffTimer = iDiff

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

Public Function ValueFromText(ByVal textin As String) As Single
' Routine Name: ValueFromText
' Created by:   Brunrose
' Function:
' This routine converts a string to a Single. If NonNumeric then zero.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim iVal As Single

SetErrModule 14, 9092
If UseLocalErrorHandler Then On Error GoTo localhandler

    If textin = Empty Then
        iVal = 0
    ElseIf textin = "" Then
        iVal = 0
    ElseIf Not IsNumeric(textin) Then
        iVal = 0
    Else
        iVal = CSng(textin)
    End If
    ValueFromText = iVal

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



