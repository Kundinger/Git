Attribute VB_Name = "Module19"
Option Explicit
'
Private rangeMax, rangeMin, span As Single

Public Sub ClearMfcCalRecords(ByVal iStation As Integer, ByVal iMfc As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Find Mfc Input Calibration Parameters
    Criteria = "SELECT * FROM [MfcCalibrations] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & "  ORDER BY [DTS] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If Not rsRecord.BOF Then
        ' record(s) found; delete the record(s)
        rsRecord.MoveFirst
        Do While Not rsRecord.BOF
            ' next cal
            rsRecord.MoveLast
'            If Not rsRecord.BOF Then dDts = rsRecord("Dts")
            ' delete cal
            If Not rsRecord.BOF Then
                dDts = rsRecord("Dts")
                rsRecord.Delete
            
                ' Find MFC Calibration Point Data
                CriteriaPts = "SELECT * FROM [MfcCalibrationsData] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
                Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
                
                ' any records found ??
                If Not rsRecordPts.BOF Then
                    ' record(s) found; delete the record(s)
                    rsRecordPts.MoveFirst
                    Do While Not rsRecordPts.BOF
                        ' next point
                        rsRecordPts.MoveLast
                        ' delete point
                        If Not rsRecordPts.BOF Then rsRecordPts.Delete
                    Loop
                End If
                ' done with points
                rsRecordPts.Close
            End If
        Loop
    End If
    
    ' done with calibrations
    rsRecord.Close
    
End Sub

Public Sub LoadMfcCalibration(ByVal iStation As Integer)
'
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim iMfc As Integer
Dim iPoint As Integer
Dim Idx As Integer
Dim tmpCal As MfcCalibration

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Read MFC Calibration Parameters
    Criteria = "SELECT * FROM [MfcCalibrations] WHERE [Station] = " & iStation & "  ORDER BY [Mfc] ASC,[Dts] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If Not rsRecord.BOF Then
        ' record(s) found; read the record(s)
        rsRecord.MoveFirst
        rsRecord.MoveLast
        Do While Not rsRecord.BOF
            iMfc = rsRecord("Mfc")
            dDts = rsRecord("Dts")
            tmpCal.dts = dDts
            tmpCal.CalibratedBy = rsRecord("CalibratedBy")
            tmpCal.Comment = rsRecord("Comment")
            tmpCal.NumPoints = rsRecord("NumPoints")
            tmpCal.RawInputType = rsRecord("RawInputType")
            tmpCal.CalData.X = rsRecord("CoefficientX")
            tmpCal.CalData.X2 = rsRecord("CoefficientX2")
            tmpCal.CalData.X3 = rsRecord("CoefficientX3")
            tmpCal.CalData.X4 = rsRecord("CoefficientX4")
            tmpCal.CalData.X5 = rsRecord("CoefficientX5")
            tmpCal.CalData.X6 = rsRecord("CoefficientX6")
            tmpCal.CalData.R2 = rsRecord("CoefficientR2")
            
            ' Read MFC Calibration Point Data
            CriteriaPts = "SELECT * FROM [MfcCalibrationsData] WHERE [Station] = " & iStation & " AND [Mfc] = " & iMfc & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
            Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
            
            ' any records found ??
            If Not rsRecordPts.BOF Then
                ' record(s) found; read the record(s)
                rsRecordPts.MoveFirst
                rsRecordPts.MoveLast
                Do While Not rsRecordPts.BOF
                    iPoint = rsRecordPts("Point")
                    tmpCal.PointData(iPoint).ActualPercent = rsRecordPts("ActualPercent")
                    tmpCal.PointData(iPoint).RawPercent = rsRecordPts("RawPercent")
                    tmpCal.PointData(iPoint).ActualValue = rsRecordPts("ActualValue")
                    tmpCal.PointData(iPoint).RawValue = rsRecordPts("RawValue")
                    ' next point
                    rsRecordPts.MovePrevious
                Loop
            End If
            ' done with points
            rsRecordPts.Close
                    
            ' copy calibration to appropriate array
            ' Station MFC Input Calibration Parameters
            Stn_MfcCal(iStation, iMfc) = tmpCal
            
            ' next input
            rsRecord.MovePrevious
        Loop
   
    End If
    ' done
    rsRecord.Close
    
End Sub

Public Sub SetMfcCalLinear(ByVal iStation As Integer)
Dim tmpVal As Single
Dim tmpval2 As Single
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim sVdcMax As Single
Dim sVdcMin As Single
Dim sVdcSpan As Single
Dim iFunc As Integer
Dim iMfc As Integer
Dim iPoint As Integer
Dim Idx As Integer
Dim tmpCal As MfcCalibration

    ' Set MFC Calibration to Linear (default)
    For iMfc = 0 To MAXMFC
            ' get analog function index for selected mfc
            Select Case iMfc
                Case MFCBUTANE
                    iFunc = asButaneFlow
                Case MFCNITROGEN
                    iFunc = asNitrogenFlow
                Case MFCPURGEAIR
                    iFunc = asPurgeAirFlow
                Case MFCORVRBUT
                    iFunc = asButaneORVRFlow
                Case MFCORVRNIT
                    iFunc = asNitrogenORVRFlow
                Case MFCORVRPRG
                    iFunc = asPurgeAirFlow
                Case MFCLIVEFUEL
                    iFunc = asLiveFuelVaporFlow
                Case MFCORVRLIVE
                    iFunc = asLiveFuelVaporORVRFlow
            End Select
            ' get min/max EU & Raw  for appropriate input
            ' Station MFC Calibration Parameters
            sEuMax = Stn_AIO(iStation, iFunc).EuMax
            sEuMin = Stn_AIO(iStation, iFunc).EuMin
            sVdcMax = Stn_AIO(iStation, iFunc).VdcMax
            sVdcMin = Stn_AIO(iStation, iFunc).VdcMin
            ' calc EU & Vdc spans
            sEuSpan = sEuMax - sEuMin
            sVdcSpan = sVdcMax - sVdcMin
            ' set calibration parameters
            tmpCal.dts = Now
            tmpCal.CalibratedBy = "default"
            tmpCal.StandardTempValue = 20
            tmpCal.StandardTempUnits = "deg C"
            tmpCal.StandardPressValue = 1
            tmpCal.StandardPressUnits = "atm"
            tmpCal.Comment = "linear"
            tmpCal.NumPoints = MAXLSQCALPOINTS
            tmpCal.RawInputType = CalRawAsVolts
'            tmpCal.CalData.X = sEuSpan
            tmpCal.CalData.X = sEuSpan
            tmpCal.CalData.X2 = CSng(0)
            tmpCal.CalData.X3 = CSng(0)
            tmpCal.CalData.X4 = CSng(0)
            tmpCal.CalData.X5 = CSng(0)
            tmpCal.CalData.X6 = CSng(0)
            tmpCal.CalData.R2 = CSng(0)
            
            ' set MFC Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                Idx = iPoint - 1
                tmpVal = CSng(Idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                tmpCal.PointData(iPoint).ActualPercent = tmpVal
                tmpCal.PointData(iPoint).RawPercent = tmpVal
                tmpCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                tmpCal.PointData(iPoint).RawValue = sVdcMin + (tmpval2 * sVdcSpan)
            Next iPoint
                    
            ' copy calibration to appropriate array
            ' Station MFC Calibration Parameters
            Stn_MfcCal(iStation, iMfc) = tmpCal
            PrevStn_MfcCal(iStation, iMfc) = tmpCal
            
    Next iMfc
    
End Sub

Public Function Cal_MfcInput(ByVal InputRaw As Single, ByVal iStation As Integer, ByVal iMfc As Integer, ByRef tmpCal As MfcCalibration) As Single
' The inputs value should be (percent full scale) / 100
' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx
' Returns the calibrated value in EngrUnits
Dim iFunc As Integer
Dim tempdbl As Double
Dim InputVal As Double
Dim xA As Double
Dim xB As Double
Dim xC As Double
Dim xD As Double
Dim xE As Double
Dim xF As Double
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim OutputEU As Single
    
    If ((iStation < 1) Or (iStation > 9)) Then
        ' output = input
        OutputEU = InputRaw
    Else
        ' get calibration coefficients
        xA = CDbl(tmpCal.CalData.X6)
        xB = CDbl(tmpCal.CalData.X5)
        xC = CDbl(tmpCal.CalData.X4)
        xD = CDbl(tmpCal.CalData.X3)
        xE = CDbl(tmpCal.CalData.X2)
        xF = CDbl(tmpCal.CalData.X)
        ' get analog function index for the selected mfc
        Select Case iMfc
            Case MFCBUTANE
                iFunc = asButaneFlow
            Case MFCNITROGEN
                iFunc = asNitrogenFlow
            Case MFCPURGEAIR
                iFunc = asPurgeAirFlow
            Case MFCORVRBUT
                iFunc = asButaneORVRFlow
            Case MFCORVRNIT
                iFunc = asNitrogenORVRFlow
            Case MFCORVRPRG
                iFunc = asPurgeAirFlow
            Case MFCLIVEFUEL
                iFunc = asLiveFuelVaporFlow
            Case MFCORVRLIVE
                iFunc = asLiveFuelVaporORVRFlow
        End Select
        ' get mfc input range values
        sEuMax = Stn_AIO(iStation, iFunc).EuMax
        sEuMin = Stn_AIO(iStation, iFunc).EuMin
        sEuSpan = sEuMax - sEuMin
        
        ' check for valid calibration
        If ((xA = 0) And (xB = 0) And (xC = 0) And (xD = 0) And (xE = 0) And (xF = 0)) Then
            ' no calibration
            ' output = linear(input)
            OutputEU = sEuMin + (InputRaw * sEuSpan)
        Else
            ' calculate calibrated value
            InputVal = CDbl(InputRaw)
            tempdbl = (CDbl(tmpCal.CalData.X) * (InputVal))
            tempdbl = tempdbl + (CDbl(tmpCal.CalData.X2) * (InputVal ^ 2))
            tempdbl = tempdbl + (CDbl(tmpCal.CalData.X3) * (InputVal ^ 3))
            tempdbl = tempdbl + (CDbl(tmpCal.CalData.X4) * (InputVal ^ 4))
            tempdbl = tempdbl + (CDbl(tmpCal.CalData.X5) * (InputVal ^ 5))
            tempdbl = tempdbl + (CDbl(tmpCal.CalData.X6) * (InputVal ^ 6))
            OutputEU = CSng(tempdbl)
        End If
    End If
    ' calibrated mfc input value
    Cal_MfcInput = OutputEU

End Function

Public Function Cal_SolveFor(ByVal DesiredVal As Single, ByVal station As Integer, ByVal MFC_No As Integer, ByVal MfcSpan As Single) As Single
    ' Returns the decimal (0.0 to 1.0) of full scale setting to use to get a desired output value
    ' Uses Newton's method to find the percent full scale value for
    ' a certain flow in SLPM
    Dim iteration_count As Integer
    Dim calx As Double
    Dim calx0 As Double
    Dim calx1 As Double
    Dim DesVal As Double
    Dim Test1 As Double
    Dim Z As Double
    Dim Tolerance_1, Tolerance_2 As Double
    
    ' Set the tolerance values
    Tolerance_1 = 0.00000000001
    Tolerance_2 = 0.00000000001
    iteration_count = 0
    
    DesVal = CDbl(DesiredVal)
    calx = DesVal / CDbl(MfcSpan)
    calx0 = calx
    calx1 = calx0

    Z = Cal_Value_D(calx0, DesVal, station, MFC_No)
    If Cal_Value_D(calx0, DesVal, station, MFC_No) <> 0 And Der_Cal_Value_D(calx0, station, MFC_No) <> 0 Then
        Do
            calx0 = calx1
            calx1 = calx0 - Cal_Value_D(calx0, DesVal, station, MFC_No) / Der_Cal_Value_D(calx0, station, MFC_No)
            iteration_count = iteration_count + 1
            If iteration_count > 18 Then
                ' We will never achieve a calibration on this value inputed
                ' What you came in with is what you leave with (NOT CALIBRATED)
                calx1 = calx
                ' Write alarm log for this process
'                If Len(StationControl(station, shift).DBFile) > 0 Then OOT_Write station, shift, "Unable to adjust MFC#" & Format(MFC_No, "0") & _
'                                                                            " calibration for DesiredValue = " & Format(DesVal, "##0.0#")
                If Len(StationControl(station, Stn_ActiveShift(station)).DBFile) > 0 Then
                    Write_ELog "Unable to calibrate AO to Station " & Format(station, "0") & " Shift " & Format(Stn_ActiveShift(station), "0") & _
                                    Mfc_Description(MFC_No) & " MFC for desired SP of " & Format(DesVal, "##0.0#") & " slpm"
                End If
                Exit Do
            End If
'            Debug.Print "iteration                                             " & Format(iteration_count, "##0")
'            Debug.Print "calx0                                                 " & Format(calx0, "###0.000000000000")
'            Debug.Print "calx1                                                 " & Format(calx1, "###0.000000000000")
'            Debug.Print "Abs(calx0 - calx1)                                    " & Format(Abs(calx0 - calx1), "##0.000000000000")
'            Debug.Print "Abs(Cal_Value_D(calx1, DesVal, station, MFC_No        " & Format(Abs(Cal_Value_D(calx1, DesVal, station, MFC_No)), "##0.000000000000")
'            Debug.Print "   "
        Loop Until (Abs(calx0 - calx1) < Tolerance_1) Or (Abs(Cal_Value_D(calx1, DesVal, station, MFC_No)) < Tolerance_2)
    End If
    Cal_SolveFor = calx1
End Function

Public Function Cal_Value_D(ByVal InputEU As Double, ByVal DesiredValue As Double, ByVal station As Integer, ByVal MFC_No As Integer) As Double
    ' Used for Newton's method calculations
    ' The inputs value should be (percent full scale) / 100
    ' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx - DesiredValue
    ' Returns the calibrated value in SLPM
    Dim tempdbl As Double
    Dim InputVal As Double
    
    InputVal = CDbl(InputEU)
    tempdbl = CDbl(Stn_MfcCal(station, MFC_No).CalData.X) * InputVal
    tempdbl = tempdbl + (CDbl(Stn_MfcCal(station, MFC_No).CalData.X2) * (InputVal ^ 2))
    tempdbl = tempdbl + (CDbl(Stn_MfcCal(station, MFC_No).CalData.X3) * (InputVal ^ 3))
    tempdbl = tempdbl + (CDbl(Stn_MfcCal(station, MFC_No).CalData.X4) * (InputVal ^ 4))
    tempdbl = tempdbl + (CDbl(Stn_MfcCal(station, MFC_No).CalData.X5) * (InputVal ^ 5))
    tempdbl = tempdbl + (CDbl(Stn_MfcCal(station, MFC_No).CalData.X6) * (InputVal ^ 6))
    tempdbl = tempdbl - DesiredValue
    Cal_Value_D = tempdbl

End Function

Public Function Der_Cal_Value_D(ByVal InputEU As Double, ByVal station As Integer, ByVal MFC_No) As Double
    ' (calx0, Station, Shift, MFC_No)
    ' This is DerCalibValue_D(), but with the Double type for the input and the output
    ' The inputs value should be (percent full scale) / 100
    
    ' Returns the derivative of the calibrated value
    ' using the formula f'(x) = 6Ax5 + 5Bx4 + 4Cx3 + 3Dx2 + 2Ex + F
    
    Dim tempdbl As Double
    Dim InputVal As Double
    
    InputVal = CDbl(InputEU)
    tempdbl = CDbl(Stn_MfcCal(station, MFC_No).CalData.X)
    tempdbl = tempdbl + (CDbl(2) * CDbl(Stn_MfcCal(station, MFC_No).CalData.X2) * (InputVal))
    tempdbl = tempdbl + (CDbl(3) * CDbl(Stn_MfcCal(station, MFC_No).CalData.X3) * (InputVal ^ 2))
    tempdbl = tempdbl + (CDbl(4) * CDbl(Stn_MfcCal(station, MFC_No).CalData.X4) * (InputVal ^ 3))
    tempdbl = tempdbl + (CDbl(5) * CDbl(Stn_MfcCal(station, MFC_No).CalData.X5) * (InputVal ^ 4))
    tempdbl = tempdbl + (CDbl(6) * CDbl(Stn_MfcCal(station, MFC_No).CalData.X6) * (InputVal ^ 5))
    
    Der_Cal_Value_D = tempdbl
   
End Function

Public Function Cal_MfcOutput(ByVal DesiredVal As Single, ByVal iStation As Integer, ByVal MFC_No As Integer, ByRef tmpCal As MfcCalibration) As Single
' Returns the decimal (0.0 to 1.0) of full scale setting to use to get a desired output value
' Uses Newton's method to find the percent full scale value for
' a certain flow in SLPM
'
Dim iFunc As Integer
' Check the full scale value for the selected MFC in the current station
Select Case MFC_No
    Case MFCBUTANE
        iFunc = asButaneFlow
    Case MFCNITROGEN
        iFunc = asNitrogenFlow
    Case MFCPURGEAIR
        iFunc = asPurgeAirFlow
    Case MFCORVRBUT
         iFunc = asButaneORVRFlow
    Case MFCORVRNIT
        iFunc = asNitrogenORVRFlow
    Case MFCORVRPRG
        iFunc = asPurgeAirFlow
    Case MFCLIVEFUEL
        iFunc = asLiveFuelVaporFlow
    Case MFCORVRLIVE
        iFunc = asLiveFuelVaporORVRFlow
End Select
rangeMax = Stn_AIO(iStation, iFunc).EuMax
rangeMin = Stn_AIO(iStation, iFunc).EuMin
span = rangeMax - rangeMin

' Avoids using regression data for uncalibrated MFC's
If tmpCal.CalData.X = 0 And tmpCal.CalData.X2 = 0 And _
    tmpCal.CalData.X3 = 0 And tmpCal.CalData.X4 = 0 And _
    tmpCal.CalData.X5 = 0 And tmpCal.CalData.X6 = 0 Then
    
    ' No Cal Data; return the Desired Value
    If (span <> 0) Then
        Cal_MfcOutput = DesiredVal / span
    Else
        Cal_MfcOutput = DesiredVal
    End If
Else
    ' ***** Cal Data Exists *****
    ' Return the value in engineering units that must be output
    ' in order to get the Desired Flow
    If (span = 0) Then
        Cal_MfcOutput = DesiredVal
    Else
        Cal_MfcOutput = Cal_SolveFor(DesiredVal, iStation, MFC_No, span)
    End If
End If
    
End Function

Public Sub ClearAiCalRecords(ByVal iGroup As Integer, ByVal iInput As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Find Analog Input Calibration Parameters
    Criteria = "SELECT * FROM [AiCalibrations] WHERE [Group] = " & iGroup & " AND [Input] = " & iInput & "  ORDER BY [DTS] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If Not rsRecord.BOF Then
        ' record(s) found; delete the record(s)
        rsRecord.MoveFirst
        Do While Not rsRecord.BOF
            ' next cal
            rsRecord.MoveLast
'            If Not rsRecord.BOF Then dDts = rsRecord("Dts")
            ' delete cal
            If Not rsRecord.BOF Then
                dDts = rsRecord("Dts")
                rsRecord.Delete
            
                ' Find Analog Input Calibration Point Data
                CriteriaPts = "SELECT * FROM [AiCalibrationsData] WHERE [Group] = " & iGroup & " AND [Input] = " & iInput & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
                Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
                
                ' any records found ??
                If Not rsRecordPts.BOF Then
                    ' record(s) found; delete the record(s)
                    rsRecordPts.MoveFirst
                    Do While Not rsRecordPts.BOF
                        ' next point
                        rsRecordPts.MoveLast
                        ' delete point
                        If Not rsRecordPts.BOF Then rsRecordPts.Delete
                    Loop
                End If
                ' done with points
                rsRecordPts.Close
            End If
        Loop
    End If
    
    ' done with calibrations
    rsRecord.Close
    
End Sub

Public Sub LoadAiCalibration(ByVal iGroup As Integer)
'
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim iInput As Integer
Dim iPoint As Integer
Dim Idx As Integer
Dim tmpCal As AICalibration

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Read Analog Input Calibration Parameters
    Criteria = "SELECT * FROM [AiCalibrations] WHERE [Group] = " & iGroup & "  ORDER BY [Input] ASC,[Dts] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If Not rsRecord.BOF Then
        ' record(s) found; read the record(s)
        rsRecord.MoveFirst
        rsRecord.MoveLast
        Do While Not rsRecord.BOF
            iInput = rsRecord("Input")
            dDts = rsRecord("Dts")
            tmpCal.dts = dDts
            tmpCal.CalibratedBy = rsRecord("CalibratedBy")
            tmpCal.Comment = rsRecord("Comment")
            tmpCal.NumPoints = rsRecord("NumPoints")
            tmpCal.RawInputType = rsRecord("RawInputType")
            tmpCal.CalData.X = rsRecord("CoefficientX")
            tmpCal.CalData.X2 = rsRecord("CoefficientX2")
            tmpCal.CalData.X3 = rsRecord("CoefficientX3")
            tmpCal.CalData.X4 = rsRecord("CoefficientX4")
            tmpCal.CalData.X5 = rsRecord("CoefficientX5")
            tmpCal.CalData.X6 = rsRecord("CoefficientX6")
            tmpCal.CalData.R2 = rsRecord("CoefficientR2")
            
            ' Read Analog Input Calibration Point Data
            CriteriaPts = "SELECT * FROM [AiCalibrationsData] WHERE [Group] = " & iGroup & " AND [Input] = " & iInput & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
            Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
            
            ' any records found ??
            If Not rsRecordPts.BOF Then
                ' record(s) found; read the record(s)
                rsRecordPts.MoveFirst
                rsRecordPts.MoveLast
                Do While Not rsRecordPts.BOF
                    iPoint = rsRecordPts("Point")
                    tmpCal.PointData(iPoint).ActualPercent = rsRecordPts("ActualPercent")
                    tmpCal.PointData(iPoint).RawPercent = rsRecordPts("RawPercent")
                    tmpCal.PointData(iPoint).ActualValue = rsRecordPts("ActualValue")
                    tmpCal.PointData(iPoint).RawValue = rsRecordPts("RawValue")
                    ' next point
                    rsRecordPts.MovePrevious
                Loop
            End If
            ' done with points
            rsRecordPts.Close
                    
            ' copy calibration to appropriate array
            Select Case iGroup
                Case calgrpComm
                    ' Common Analog Input Calibration Parameters
                    Com_AiCal(iInput) = tmpCal
                    
                Case calgrpStn1 To calgrpStn9
                    ' Station Analog Input Calibration Parameters
                    Stn_AiCal(iGroup, iInput) = tmpCal
                    
                Case calgrpFid
                    ' FID Analog Input Calibration Parameters
'                    Fid_AiCal(iInput) = tmpCal
                    
                Case calgrpPrg1 To calgrpPrg9
                    ' Purge Analog Input Calibration Parameters
                    Idx = iGroup - 10
                    Prg_AiCal(Idx, iInput) = tmpCal
                    
            End Select
            
            ' next input
            rsRecord.MovePrevious
        Loop
   
    End If
    ' done
    rsRecord.Close
    
End Sub

Public Sub SetAiCalLinear(ByVal iGroup As Integer)
Dim tmpVal As Single
Dim tmpval2 As Single
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim sVdcMax As Single
Dim sVdcMin As Single
Dim sVdcSpan As Single
Dim iMax As Integer
Dim iInput As Integer
Dim iPoint As Integer
Dim Idx As Integer
Dim tmpCal As AICalibration

    ' get max input index for appropriate array
    Select Case iGroup
        Case calgrpComm
            ' Common Analog Input Calibration Parameters
            iMax = MAX_ANA_COM
        Case calgrpStn1 To calgrpStn9
            ' Station Analog Input Calibration Parameters
            iMax = MAX_ANA_STN
        Case calgrpFid
            ' FID Analog Input Calibration Parameters
'            iMax = MAX_ANA_FID
        Case calgrpPrg1 To calgrpPrg9
            ' Purge Analog Input Calibration Parameters
            iMax = MAX_ANA_PRG
    End Select
            
    ' Set Analog Input Calibration to Linear (default)
    For iInput = 1 To iMax
            ' get min/max EU & Raw  for appropriate input
            Select Case iGroup
                Case calgrpComm
                    ' Common Analog Input Calibration Parameters
                    sEuMax = Com_AIO(iInput).EuMax
                    sEuMin = Com_AIO(iInput).EuMin
                    sVdcMax = Com_AIO(iInput).VdcMax
                    sVdcMin = Com_AIO(iInput).VdcMin
                Case calgrpStn1 To calgrpStn9
                    ' Station Analog Input Calibration Parameters
                    sEuMax = Stn_AIO(iGroup, iInput).EuMax
                    sEuMin = Stn_AIO(iGroup, iInput).EuMin
                    sVdcMax = Stn_AIO(iGroup, iInput).VdcMax
                    sVdcMin = Stn_AIO(iGroup, iInput).VdcMin
                Case calgrpFid
                    ' FID Analog Input Calibration Parameters
'                    sEuMax = Fid_AIO(iInput).EuMax
'                    sEuMin = Fid_AIO(iInput).EuMin
'                    sVdcMax = Fid_AIO(iInput).VdcMax
'                    sVdcMin = Fid_AIO(iInput).VdcMin
                Case calgrpPrg1 To calgrpPrg9
                    ' Purge Analog Input Calibration Parameters
                    Idx = iGroup - 10
                    sEuMax = Prg_AIO(Idx, iInput).EuMax
                    sEuMin = Prg_AIO(Idx, iInput).EuMin
                    sVdcMax = Prg_AIO(Idx, iInput).VdcMax
                    sVdcMin = Prg_AIO(Idx, iInput).VdcMin
            End Select
            ' calc EU & Vdc spans
            sEuSpan = sEuMax - sEuMin
            sVdcSpan = sVdcMax - sVdcMin
            ' set calibration parameters
            tmpCal.dts = Now
            tmpCal.CalibratedBy = "default"
            tmpCal.StandardTempValue = 20
            tmpCal.StandardTempUnits = "deg C"
            tmpCal.StandardPressValue = 1
            tmpCal.StandardPressUnits = "atm"
            tmpCal.Comment = "linear"
            tmpCal.NumPoints = MAXLSQCALPOINTS
            tmpCal.RawInputType = CalRawAsVolts
'            tmpCal.CalData.X = sEuSpan
            tmpCal.CalData.X = CSng(0)
            tmpCal.CalData.X2 = CSng(0)
            tmpCal.CalData.X3 = CSng(0)
            tmpCal.CalData.X4 = CSng(0)
            tmpCal.CalData.X5 = CSng(0)
            tmpCal.CalData.X6 = CSng(0)
            tmpCal.CalData.R2 = CSng(0)
            
            ' set Analog Input Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                Idx = iPoint - 1
                tmpVal = CSng(Idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                tmpCal.PointData(iPoint).ActualPercent = tmpVal
                tmpCal.PointData(iPoint).RawPercent = tmpVal
                tmpCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                tmpCal.PointData(iPoint).RawValue = sVdcMin + (tmpval2 * sVdcSpan)
            Next iPoint
                    
            ' copy calibration to appropriate array
            Select Case iGroup
                Case calgrpComm
                    ' Common Analog Input Calibration Parameters
                    Com_AiCal(iInput) = tmpCal
                    PrevCom_AiCal(iInput) = tmpCal
                Case calgrpStn1 To calgrpStn9
                    ' Station Analog Input Calibration Parameters
                    Stn_AiCal(iGroup, iInput) = tmpCal
                    PrevStn_AiCal(iGroup, iInput) = tmpCal
                Case calgrpFid
                    ' FID Analog Input Calibration Parameters
'                    Fid_AiCal(iInput) = tmpCal
'                    PrevFid_AiCal(iInput) = tmpCal
                Case calgrpPrg1 To calgrpPrg9
                    ' Purge Analog Input Calibration Parameters
                    Idx = iGroup - 10
                    Prg_AiCal(Idx, iInput) = tmpCal
                    PrevPrg_AiCal(Idx, iInput) = tmpCal
            End Select
            
    Next iInput
    
End Sub

Public Function Cal_AnalogInput(ByVal InputRaw As Single, ByVal iGroup As Integer, ByVal iInput As Integer, ByRef tmpCal As AICalibration) As Single
' The inputs value should be (percent full scale) / 100
' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx
' Returns the calibrated value in EngrUnits
Dim Idx As Integer
Dim tempdbl As Double
Dim InputVal As Double
Dim xA As Double
Dim xB As Double
Dim xC As Double
Dim xD As Double
Dim xE As Double
Dim xF As Double
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim OutputEU As Single
    
    If ((iGroup < calgrpComm) Or (iGroup > (MAX_PRG + 10))) Then
        ' output = input
        OutputEU = InputRaw
    Else
        ' get calibration coefficients
        xA = CDbl(tmpCal.CalData.X6)
        xB = CDbl(tmpCal.CalData.X5)
        xC = CDbl(tmpCal.CalData.X4)
        xD = CDbl(tmpCal.CalData.X3)
        xE = CDbl(tmpCal.CalData.X2)
        xF = CDbl(tmpCal.CalData.X)
        ' get Analog Input Range values
        Select Case iGroup
            Case calgrpComm
                sEuMax = Com_AIO(iInput).EuMax
                sEuMin = Com_AIO(iInput).EuMin
            Case calgrpStn1 To calgrpStn9
                Idx = iGroup
                sEuMax = Stn_AIO(Idx, iInput).EuMax
                sEuMin = Stn_AIO(Idx, iInput).EuMin
            Case calgrpFid
'                sEuMax = Fid_AIO(iInput).EuMax
'                sEuMin = Fid_AIO(iInput).EuMin
            Case calgrpPrg1 To calgrpPrg9
                Idx = iGroup - 10
                sEuMax = Prg_AIO(Idx, iInput).EuMax
                sEuMin = Prg_AIO(Idx, iInput).EuMin
        End Select
        sEuSpan = sEuMax - sEuMin
        
        ' check for valid calibration
        If ((xA = 0) And (xB = 0) And (xC = 0) And (xD = 0) And (xE = 0) And (xF = 0)) Then
            ' no calibration
            ' output = linear(input)
            OutputEU = sEuMin + (InputRaw * sEuSpan)
        Else
            ' calculate calibrated value
            InputVal = CDbl(InputRaw)
            tempdbl = (xF * (InputVal))
            tempdbl = tempdbl + (xE * (InputVal ^ 2))
            tempdbl = tempdbl + (xD * (InputVal ^ 3))
            tempdbl = tempdbl + (xC * (InputVal ^ 4))
            tempdbl = tempdbl + (xB * (InputVal ^ 5))
            tempdbl = tempdbl + (xA * (InputVal ^ 6))
            OutputEU = sEuMin + (CSng(tempdbl) * sEuSpan)
        End If
    End If
    ' calibrated analog input value
    Cal_AnalogInput = OutputEU

End Function

Public Sub ClearSclCalRecords(ByVal iScale As Integer)
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' Find Analog Input Calibration Parameters
    Criteria = "SELECT * FROM [ScaleCalibrations] WHERE [Scale] = " & iScale & "  ORDER BY [DTS] DESC"
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' any records found ??
    If Not rsRecord.BOF Then
        ' record(s) found; delete the record(s)
        rsRecord.MoveFirst
        Do While Not rsRecord.BOF
            ' next cal
            rsRecord.MoveLast
'            If Not rsRecord.BOF Then dDts = rsRecord("Dts")
            ' delete cal
            If Not rsRecord.BOF Then
                dDts = rsRecord("Dts")
                rsRecord.Delete
            
                ' Find Scale Calibration Point Data
                CriteriaPts = "SELECT * FROM [ScaleCalibrationsData] WHERE [Scale] = " & iScale & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
                Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
                
                ' any records found ??
                If Not rsRecordPts.BOF Then
                    ' record(s) found; delete the record(s)
                    rsRecordPts.MoveFirst
                    Do While Not rsRecordPts.BOF
                        ' next point
                        rsRecordPts.MoveLast
                        ' delete point
                        If Not rsRecordPts.BOF Then rsRecordPts.Delete
                    Loop
                End If
                ' done with points
                rsRecordPts.Close
            End If
        Loop
    End If
    
    ' done with calibrations
    rsRecord.Close
    
End Sub

Public Sub LoadSclCalibration()
Dim dbDbase As Database
Dim rsRecord  As Recordset
Dim Criteria As String
Dim dbDbasePts As Database
Dim rsRecordPts  As Recordset
Dim CriteriaPts As String
Dim dDts As Date
Dim flagDone As Boolean
Dim iScale As Integer
Dim iPoint As Integer
Dim CurrPrev As String
Dim tmpCal As SclCalibration

    ' open calibration database
    Set dbDbase = OpenDatabase(FILEPATH_cal & DATACAL)
    Set dbDbasePts = OpenDatabase(FILEPATH_cal & DATACAL)
    
    ' cycle through all scales
    For iScale = 1 To NR_SCALES
    
        flagDone = False
        CurrPrev = "Current"
    
        ' Read Scale Calibration Parameters
        Criteria = "SELECT * FROM [ScaleCalibrations] WHERE [Scale] = " & iScale & " ORDER BY [Dts] ASC"
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        
        ' any records found ??
        If Not rsRecord.BOF Then
            ' record(s) found; read the record(s)
            rsRecord.MoveFirst
            rsRecord.MoveLast
            Do While Not rsRecord.BOF And Not flagDone
                dDts = rsRecord("Dts")
                tmpCal.dts = dDts
                tmpCal.CalibratedBy = rsRecord("CalibratedBy")
                tmpCal.Comment = rsRecord("Comment")
                tmpCal.CalRangeMax = rsRecord("CalRangeMax")
                tmpCal.CalRangeMin = rsRecord("CalRangeMin")
                tmpCal.NumPoints = rsRecord("NumPoints")
                tmpCal.CalData.X = rsRecord("CoefficientX")
                tmpCal.CalData.X2 = rsRecord("CoefficientX2")
                tmpCal.CalData.X3 = rsRecord("CoefficientX3")
                tmpCal.CalData.X4 = rsRecord("CoefficientX4")
                tmpCal.CalData.X5 = rsRecord("CoefficientX5")
                tmpCal.CalData.X6 = rsRecord("CoefficientX6")
                tmpCal.CalData.R2 = rsRecord("CoefficientR2")
                
                ' CALIBRATION POINT DATA
                ' clear scale calibration point data
                For iPoint = MINLSQCALPOINTS To MAXLSQCALPOINTS
                    tmpCal.PointData(iPoint).ActualPercent = 0
                    tmpCal.PointData(iPoint).RawPercent = 0
                    tmpCal.PointData(iPoint).ActualValue = 0
                    tmpCal.PointData(iPoint).RawValue = 0
                Next iPoint
                ' Read Scale Calibration Point Data
                CriteriaPts = "SELECT * FROM [ScaleCalibrationsData] WHERE [Scale] = " & iScale & " AND [DTS] = #" & dDts & "#  ORDER BY [Point] ASC"
                Set rsRecordPts = dbDbasePts.OpenRecordset(CriteriaPts, dbOpenDynaset)
                ' any records found ??
                If Not rsRecordPts.BOF Then
                    ' record(s) found; read the record(s)
                    rsRecordPts.MoveFirst
                    rsRecordPts.MoveLast
                    Do While Not rsRecordPts.BOF
                        iPoint = rsRecordPts("Point")
                        tmpCal.PointData(iPoint).ActualPercent = rsRecordPts("ActualPercent")
                        tmpCal.PointData(iPoint).RawPercent = rsRecordPts("RawPercent")
                        tmpCal.PointData(iPoint).ActualValue = rsRecordPts("ActualValue")
                        tmpCal.PointData(iPoint).RawValue = rsRecordPts("RawValue")
                        ' next point
                        rsRecordPts.MovePrevious
                    Loop
                End If
                ' done with points
                rsRecordPts.Close
                        
                ' copy calibration to appropriate scale
                Select Case CurrPrev
                    Case "Current"
                        Scale_Cal(iScale) = tmpCal
                    Case "Previous"
                        PrevScale_Cal(iScale) = tmpCal
                        flagDone = True
                    Case Else
                End Select
                CurrPrev = "Previous"
                ' previous record
                rsRecord.MovePrevious
            Loop
        
        End If
        ' done
        rsRecord.Close
        
    Next iScale
     
End Sub

Public Sub SetSclCalLinear()
Dim tmpVal As Single
Dim tmpval2 As Single
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim Idx As Integer
Dim iScale As Integer
Dim iPoint As Integer
Dim tmpCal As SclCalibration

    ' Set Scale Calibration to Linear (default)
    For iScale = 1 To MAX_SCALES
            ' set calibration parameters
            tmpCal.dts = Now
            tmpCal.StandardTempValue = 20
            tmpCal.StandardTempUnits = "deg C"
            tmpCal.StandardPressValue = 1
            tmpCal.StandardPressUnits = "atm"
            tmpCal.CalibratedBy = "clear"
            tmpCal.Comment = "na"
            tmpCal.NumPoints = MAXLSQCALPOINTS
            tmpCal.CalRangeMax = CSng(DefScaleMax)
            tmpCal.CalRangeMin = CSng(0)
            tmpCal.CalData.X = CSng(0)
            tmpCal.CalData.X2 = CSng(0)
            tmpCal.CalData.X3 = CSng(0)
            tmpCal.CalData.X4 = CSng(0)
            tmpCal.CalData.X5 = CSng(0)
            tmpCal.CalData.X6 = CSng(0)
            tmpCal.CalData.R2 = CSng(0)
            ' calc EU span
            sEuMax = tmpCal.CalRangeMax
            sEuMin = tmpCal.CalRangeMin
            sEuSpan = sEuMax - CSng(0)
            
            ' set Scale Calibration Point Data
            For iPoint = 1 To MAXLSQCALPOINTS
                Idx = iPoint - 1
                tmpVal = CSng(Idx) * CSng(10)
                tmpval2 = tmpVal / CSng(100)
                tmpCal.PointData(iPoint).ActualPercent = tmpVal
                tmpCal.PointData(iPoint).RawPercent = tmpVal
                tmpCal.PointData(iPoint).ActualValue = sEuMin + (tmpval2 * sEuSpan)
                tmpCal.PointData(iPoint).RawValue = sEuMin + (tmpval2 * sEuSpan)
            Next iPoint
                    
            ' copy calibration to appropriate scale
            Scale_Cal(iScale) = tmpCal
            PrevScale_Cal(iScale) = tmpCal
            
    Next iScale
    
End Sub

Public Function Cal_Scale(ByVal InputRaw As Single, ByVal iScale As Integer, ByRef tmpCal As SclCalibration) As Single
' The inputs value should be (percent full scale) / 100
' Uses the formula f(x) = Ax6 + Bx5 + Cx4 + Dx3 + Ex2 + Fx
' Returns the calibrated value in EngrUnits
Dim Idx As Integer
Dim tempdbl As Double
Dim InputVal As Double
Dim xA As Double
Dim xB As Double
Dim xC As Double
Dim xD As Double
Dim xE As Double
Dim xF As Double
Dim normVal As Single
Dim sEuMax As Single
Dim sEuMin As Single
Dim sEuSpan As Single
Dim OutputEU As Single
    
    ' get calibration coefficients
    xA = CDbl(tmpCal.CalData.X6)
    xB = CDbl(tmpCal.CalData.X5)
    xC = CDbl(tmpCal.CalData.X4)
    xD = CDbl(tmpCal.CalData.X3)
    xE = CDbl(tmpCal.CalData.X2)
    xF = CDbl(tmpCal.CalData.X)
    ' get scale span value
    sEuMax = tmpCal.CalRangeMax
    sEuMin = tmpCal.CalRangeMin
    sEuSpan = sEuMax - sEuMin
    
    ' check for valid calibration
    If ((xA = 0) And (xB = 0) And (xC = 0) And (xD = 0) And (xE = 0) And (xF = 0)) Then
        ' no calibration
        ' output = linear(input)
        OutputEU = InputRaw
    Else
        ' calculate calibrated value
        normVal = (InputRaw - sEuMin) / sEuSpan
        InputVal = CDbl(InputRaw)
        tempdbl = (xF * (InputVal))
        tempdbl = tempdbl + (xE * (InputVal ^ 2))
        tempdbl = tempdbl + (xD * (InputVal ^ 3))
        tempdbl = tempdbl + (xC * (InputVal ^ 4))
        tempdbl = tempdbl + (xB * (InputVal ^ 5))
        tempdbl = tempdbl + (xA * (InputVal ^ 6))
'        OutputEU = CSng(tempdbl)
        OutputEU = sEuMin + (CSng(tempdbl) * sEuSpan)
'        OutputEU = InputRaw
    End If
    ' calibrated scale value
    Cal_Scale = OutputEU

End Function


