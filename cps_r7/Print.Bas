Attribute VB_Name = "Module10"
'error module 10 '''''''''''''' program PRINT.bas '''''''''''''''''''
' Printing Functions Module
Option Explicit

Sub File_Center(ByVal iFile As Integer, ByVal fString As String)
'
' Procedure Name:   File_Center
' Created by:       Analytical Process Programmer 9/96
' Description:      This routine sends a line of text to the specified file
'                   and centers the line based upon the current files
'                   font and margins.
'
'

' Adds necessary spacing to center a line for writing to the ASCII
' formated file.  If the length of the string provided is too long
' it is written without modification


If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 11

Dim icnt As Integer
Dim USpace As Integer

    ' Get Usable page space
    USpace = FILEPAGECOLS - RTMARGIN - LTMARGIN
    If Len(fString) < USpace Then
      icnt = CInt((USpace - Len(fString)) / 2)
      fString = String(icnt, " ") & fString
    End If
    ' write the line to the file
    File_Line iFile, fString

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

Sub File_Line(ByVal iFile As Integer, ByVal fString As String)
'
' Function Name:    File_Line
' Author            Analytical Process Programmer 9/96
' Description:      Sends one line to an ASCII formatted file.
'                   The file handle is provided by iFile, the string to
'                   send is fstring.
'
'                   Maintains ASCII file formatting for 80 column, 66 line
'                   sheet.  Accounts for Top and Bottom margins and pages
'                   to new page when required.  Page footer and form feed
'                   is provided by footer routine.
'
'

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 13

Dim spaceleft As Integer
Dim maxsize As Integer
Dim pString As String
Dim substring As String
Dim bDone As Boolean

    ' set maximum string size as number of cols, including left margin
    maxsize = FILEPAGECOLS - RTMARGIN - 1
    bDone = False
    ' Print only max characters allowed per line
    ' Print new lines if exceeding max characters.
    Do While Not (bDone)
     ' Add left margin spacing
     fString = String(LTMARGIN, " ") & fString
     ' If the line contains any carraige return, line feeds, get rid of them
     ' the Print # contains them
     If InStr(fString, vbCrLf) Then
       If Len(fString) > InStr(fString, vbCrLf) + 1 Then
         substring = Left(fString, InStr(fString, vbCrLf) - 1)
         fString = Mid$(fString, InStr(fString, vbCrLf) + 2)
       Else
         substring = Left(fString, InStr(fString, vbCrLf) - 1)
         fString = ""
       End If
       If Len(substring) > maxsize Then
         pString = Left(substring, maxsize)
         fString = Mid$(substring, maxsize + 1) & vbCrLf & fString
       Else
         pString = substring
       End If
     Else  ' No Carraige Returns
       If Len(fString) > maxsize Then
         pString = Left(fString, maxsize)
         fString = Mid$(fString, maxsize + 1)
       Else
         pString = fString
         bDone = True
       End If
     End If
     Print #iFile, pString
    Loop

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

Sub Print_File(ByVal sFileName As String)
'
' Procedure Name:   Print_File
' Created By:       Analytical Process Programmer 9/96
' Description:
' This procedure prints an ASCII formatted text file to the printer
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 2

Dim iFileNumber As Integer
Dim linetoprint As String
Dim oldFont As New StdFont

    ' Save current printer font
    oldFont = Printer.Font
    Printer.Font = FILEFONT
    Printer.Font.Size = FILEFONTSIZE
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    ' Sometimes get error if in page, no need to alter this here
    'Printer.Orientation = 1 ' portrait
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    Do While Not EOF(iFileNumber)
        Line Input #iFileNumber, linetoprint
        Print_Line linetoprint
    Loop
    Close #iFileNumber
    Print_Footer
    Printer.EndDoc
    Printer.Font = oldFont
    
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

Sub Print_Center(linetoprint As String)
'
' Procedure Name:   Print_Center
' Created by:       Analytical Process Programmer 9/96
' Description:      This routine sets the x position for the input string
'                   and calls the print_line routine to print the current
'                   line.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 1
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(linetoprint)) / 2
    Print_Line linetoprint
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Sub Print_Footer(Optional ByVal numPages As Integer)
'
' Procedure Name:   Print_Footer
' Created by:       Analytical Process Programmer 9/96
' Description:      This routine prints a footer with page number to the
'                   current printer object.
'                   Print_footer should be followed by a newpage command
'
Dim footerText As String

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 3
    footerText = "Page " & Format$(Printer.Page, "###")
    If numPages > 0 Then footerText = footerText & " of " & Format$(numPages, "###")
    Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(footerText)) / 2
    Printer.CurrentY = Printer.ScaleHeight - Printer.TextHeight(footerText)
    Printer.Print footerText
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Sub Print_Line(linetoprint As String)
'
' Procedure Name:   PrintLine
' Created by:       Analytical Process Programmer     9/96
' Description:      This routine accepts a single line of input and prints
'                   it to the printer object.  If the number of lines
'                   remaining is less than 4, a form feed is sent and
'                   the page number is printed on the bottom of the form.
'
'                   The routine preserves the current x position to allow
'                   centering routines to work.
'
Dim xposition As Integer
Dim spaceleft As Integer
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 10, 4
    xposition = Printer.CurrentX
    ' Check to see if below top margin
    While Printer.CurrentY < (TOPMARGIN * Printer.TextHeight("TEST"))
      Printer.Print
    Wend
    ' Determine remaining space on form
    spaceleft = Printer.ScaleHeight - Printer.CurrentY
    ' See if below bottom margin
    If spaceleft < BOTTOMMARGIN * Printer.TextHeight(linetoprint) Then
      Print_Footer
      Printer.NewPage
      ' Add spaces till past top margin
      While Printer.CurrentY < (TOPMARGIN * Printer.TextHeight("TEST"))
        Printer.Print
      Wend
    End If
    Printer.CurrentX = xposition
    Printer.Print linetoprint
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



