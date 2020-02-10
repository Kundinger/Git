VERSION 5.00
Begin VB.Form frmPrintSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Setup"
   ClientHeight    =   4455
   ClientLeft      =   2175
   ClientTop       =   1440
   ClientWidth     =   9375
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   9375
   Begin VB.Frame frmSamples 
      Caption         =   "Samples"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5880
      TabIndex        =   18
      Top             =   2640
      Width           =   1695
      Begin VB.Label lblFileFontSample 
         Alignment       =   2  'Center
         Caption         =   "AaBbYyZz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblReportFontSample 
         Alignment       =   2  'Center
         Caption         =   "AaBbYyZz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   870
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "      Restore     Default Fonts"
      DisabledPicture =   "frmPrintSet.frx":57E2
      DownPicture     =   "frmPrintSet.frx":5EE4
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7785
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrintSet.frx":65E6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
   Begin VB.Frame frmFonts 
      Caption         =   "Fonts"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   5415
      Begin VB.ComboBox SelectReportFont 
         Height          =   315
         Left            =   2400
         TabIndex        =   16
         Text            =   "SelectReportFont"
         Top             =   840
         Width           =   2775
      End
      Begin VB.ComboBox SelectFileFont 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Text            =   "SelectFileFont"
         Top             =   450
         Width           =   2775
      End
      Begin VB.Label lblReportFont 
         Caption         =   " Other Reports:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label lblFileFont 
         Caption         =   " Detail and Summary Reports:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Printer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8895
      Begin VB.ComboBox SelectPrinter 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   330
         Width           =   7365
      End
      Begin VB.Label lblComment 
         Caption         =   " Comment:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblCommentDesc 
         Caption         =   "comment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   10
         Top             =   1560
         Width           =   7365
      End
      Begin VB.Label lblLocation 
         Caption         =   " Where:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lblLocationDesc 
         Caption         =   "location"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   1320
         Width           =   7365
      End
      Begin VB.Label lblType 
         Caption         =   " Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblTypeDesc 
         Caption         =   "type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   1080
         Width           =   7365
      End
      Begin VB.Label lblStatusDesc 
         Caption         =   "status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   7365
      End
      Begin VB.Label lblStatus 
         Caption         =   " Status:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblName 
         Caption         =   " Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      DisabledPicture =   "frmPrintSet.frx":6CE8
      DownPicture     =   "frmPrintSet.frx":792A
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7785
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPrintSet.frx":856C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   1350
   End
End
Attribute VB_Name = "frmPrintSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 94 ''''''''''''' Form PRINTSET.frm ''''''''''''''''''''''''
Option Explicit
Private Declare Function GetTextMetrics Lib "gdi32" _
   Alias "GetTextMetricsA" _
   (ByVal hdc As Long, _
   lpMetrics As TEXTMETRIC) _
   As Long

Private Type TEXTMETRIC
   tmHeight As Long
   tmAscent As Long
   tmDescent As Long
   tmInternalLeading As Long
   tmExternalLeading As Long
   tmAveCharWidth As Long
   tmMaxCharWidth As Long
   tmWeight As Long
   tmOverhang As Long
   tmDigitizedAspectX As Long
   tmDigitizedAspectY As Long
   tmFirstChar As Byte
   tmLastChar As Byte
   tmDefaultChar As Byte
   tmBreakChar As Byte
   tmItalic As Byte
   tmUnderlined As Byte
   tmStruckOut As Byte
   tmPitchAndFamily As Byte
   tmCharSet As Byte
End Type

Const FIXED_PITCH_BIT As Byte = 1

Private fileFontList() As New StdFont
Private reportFontList() As New StdFont


Private Sub cmdClose_Click()
SetErrModule 94, 1
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    Unload Me
    Set frmPrintSet = Nothing
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub FindFileFonts()
Dim a, b, c, curr As Integer
Dim Index, index2 As Long
Dim tm As TEXTMETRIC
Dim ret As Long
Dim FontFound As Boolean
Dim tmpfont, prevFont As New StdFont
ReDim fileFontList(0)

SetErrModule 94, 3
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    prevFont = Printer.Font
    FontFound = False   ' Just in case none are found!
    SelectFileFont.Clear
    c = 0
    For Index = 0 To Printer.FontCount - 1      ' Determine number of fonts.
       Printer.FontName = Printer.Fonts(Index)  ' Select the font.
       ret = GetTextMetrics(Printer.hdc, tm)    ' Retrieve information.
       ' Test the fixed pitch bit.
       ' Fonts with this bit off are fixed pitch.
       If (tm.tmPitchAndFamily And FIXED_PITCH_BIT) = 0 Then
           If Mid(Printer.FontName, 1, 1) <> "@" Then
               ReDim Preserve fileFontList(UBound(fileFontList) + 1)
               fileFontList(UBound(fileFontList)) = Printer.FontName
               FontFound = True   ' Found at least one!
               c = c + 1
           End If
       End If
    Next Index
    If Not FontFound Then
       SelectFileFont.AddItem "No fixed pitched fonts found!", 0
    Else
        ' Bubble sort the fonts
        For a = 0 To UBound(fileFontList) - 1
            For b = a + 1 To UBound(fileFontList)
                If fileFontList(a).Name > fileFontList(b).Name Then
                    tmpfont = fileFontList(a)
                    fileFontList(a) = fileFontList(b)
                    fileFontList(b) = tmpfont
                End If
            Next b
        Next a
        ' add fonts to combo
        c = 0
        curr = 0
        For index2 = 0 To UBound(fileFontList)
            If index2 = 0 Then
                SelectFileFont.AddItem fileFontList(index2).Name, c
                c = c + 1
            ElseIf fileFontList(index2).Name <> fileFontList(index2 - 1).Name Then
                SelectFileFont.AddItem fileFontList(index2).Name, c
                If fileFontList(index2).Name = FILEFONT.Name Then curr = c
                c = c + 1
            End If
        Next index2
        SelectFileFont.ListIndex = curr
    End If
    Printer.Font = prevFont
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub FindReportFonts()
Dim a, b, c, curr As Integer
Dim Index, index2 As Long
Dim tm As TEXTMETRIC
Dim ret As Long
Dim FontFound As Boolean
Dim tmpfont, prevFont As New StdFont
ReDim reportFontList(0)

SetErrModule 94, 4
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    prevFont = Printer.Font
    FontFound = False                           ' Just in case none are found!
    SelectReportFont.Clear
    c = 0
    For Index = 0 To Printer.FontCount - 1      ' Determine number of fonts.
       Printer.FontName = Printer.Fonts(Index)  ' Select the font.
       ret = GetTextMetrics(Printer.hdc, tm)    ' Retrieve information.
       ' Test the fixed pitch bit.
       ' Fonts with this bit off are fixed pitch.
       If (tm.tmPitchAndFamily And FIXED_PITCH_BIT) <> 0 Then
           If Mid(Printer.FontName, 1, 1) <> "@" Then
                ReDim Preserve reportFontList(UBound(reportFontList) + 1)
                reportFontList(UBound(reportFontList)) = Printer.FontName
                FontFound = True   ' Found at least one!
                c = c + 1
           End If
       End If
    Next Index
    If Not FontFound Then
       SelectReportFont.AddItem "No fonts found!", Index
    Else
        ' Bubble sort the fonts
        For a = 0 To UBound(reportFontList) - 1
            For b = a + 1 To UBound(reportFontList)
                If reportFontList(a).Name > reportFontList(b).Name Then
                    tmpfont = reportFontList(a)
                    reportFontList(a) = reportFontList(b)
                    reportFontList(b) = tmpfont
                End If
            Next b
        Next a
        ' add fonts to combo
        c = 0
        curr = 0
        For index2 = 0 To UBound(reportFontList)
            If index2 = 0 Then
                SelectReportFont.AddItem reportFontList(index2).Name, c
                c = c + 1
            ElseIf reportFontList(index2).Name <> reportFontList(index2 - 1).Name Then
                SelectReportFont.AddItem reportFontList(index2).Name, c
                If reportFontList(index2).Name = REPORTFONT.Name Then curr = c
                c = c + 1
            End If
        Next index2
        SelectReportFont.ListIndex = curr
    End If
    Printer.Font = prevFont

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub cmdDefaults_Click()
Dim def, Index As Integer

SetErrModule 94, 3
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    ' find default File Font index
    def = 0
    For Index = 0 To UBound(fileFontList)
        If fileFontList(Index).Name = "Lucida Console" Then
            def = Index
        End If
    Next Index
    ' restore File Font
    SelectFileFont.ListIndex = def
    ' reset sample characteristics
    lblFileFontSample.FontSize = 10
    lblFileFontSample.FontBold = False
    lblFileFontSample.FontItalic = False
    lblFileFontSample.FontStrikethru = False
    lblFileFontSample.FontUnderline = False
    lblFileFontSample.ForeColor = BLACK
    
    ' find default Report Font index
    def = 0
    For Index = 0 To UBound(reportFontList)
        If reportFontList(Index).Name = "Arial" Then
            def = Index
        End If
    Next Index
    ' restore Report Font
    SelectReportFont.ListIndex = def
    ' reset sample characteristics
    lblReportFontSample.FontSize = 10
    lblReportFontSample.FontBold = False
    lblReportFontSample.FontItalic = False
    lblReportFontSample.FontStrikethru = False
    lblReportFontSample.FontUnderline = False
    lblReportFontSample.ForeColor = BLACK

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub Form_Load()
Dim X As Printer
Dim idx, curr As Integer

SetErrModule 94, 0
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    Form_Center Me

    FindFileFonts
    FindReportFonts

    idx = 0
    For Each X In Printers
        SelectPrinter.AddItem X.DeviceName, idx
        If PrinterName = X.DeviceName Then
            ' this printer is the current Printer
            curr = idx
        End If
        idx = idx + 1
    Next X
    SelectPrinter.ListIndex = curr

'Printers (Index)
'The index placeholder represents an integer with a range from 0 to Printers.Count-1.

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmPrintSet = Nothing     'current form
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub SelectFileFont_Click()
SetErrModule 94, 5
If UseLocalErrorHandler Then On Error GoTo localhandler

    FILEFONT.Name = SelectFileFont.List(SelectFileFont.ListIndex)
    ' reset sample characteristics
    lblFileFontSample.FontSize = 10
    lblFileFontSample.FontBold = False
    lblFileFontSample.FontItalic = False
    lblFileFontSample.FontStrikethru = False
    lblFileFontSample.FontUnderline = False
    lblFileFontSample.ForeColor = BLACK
    ' update sample
    lblFileFontSample.Font = FILEFONT
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub SelectReportFont_Click()
SetErrModule 94, 6
If UseLocalErrorHandler Then On Error GoTo localhandler

    REPORTFONT.Name = SelectReportFont.List(SelectReportFont.ListIndex)
    ' reset sample characteristics
    lblReportFontSample.FontBold = False
    lblReportFontSample.FontItalic = False
    lblReportFontSample.FontStrikethru = False
    lblReportFontSample.FontUnderline = False
    lblReportFontSample.ForeColor = BLACK
    ' update sample
    lblReportFontSample.Font = REPORTFONT
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub SelectPrinter_Click()
SetErrModule 94, 7
If UseLocalErrorHandler Then On Error GoTo localhandler

    PrinterName = Printers(SelectPrinter.ListIndex).DeviceName
    Set Printer = Printers(SelectPrinter.ListIndex)
    CheckPrinter
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Sub CheckPrinter()
Dim hPrinter As Long
Dim ByteBuf As Long
Dim BytesNeeded As Long
Dim PI2 As PRINTER_INFO_2
Dim PrinterInfo() As Byte
Dim Result As Long
Dim LastError As Long
Dim pDefaults As PRINTER_DEFAULTS

SetErrModule 94, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

    'Set desired access security setting.
    pDefaults.DesiredAccess = PRINTER_ACCESS_USE
    
    'Call API to get a handle to the printer.
    Result = OpenPrinter(PrinterName, hPrinter, pDefaults)
    If Result = 0 Then
       'If an error occurred, display an error and exit sub.
       lblCommentDesc.Caption = "Cannot open printer " & PrinterName & _
          ", Error: " & err.LastDllError
       Exit Sub
    End If
    
    'Init BytesNeeded
    BytesNeeded = 0
    
    'Clear the error object of any errors.
    err.Clear
    
    'Determine the buffer size that is needed to get printer info.
    Result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
    
    'Check for error calling GetPrinter.
    If err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
       'Display an error message, close printer, and exit sub.
       lblCommentDesc.Caption = " > GetPrinter Failed on initial call! <"
       ClosePrinter hPrinter
       Exit Sub
    End If
    
    'Note that in Charles Petzold's book "Programming Windows 95," he
    'states that because of a problem with GetPrinter on Windows 95 only, you
    'must allocate a buffer as much as three times larger than the value
    'returned by the initial call to GetPrinter. This is not done here.
    ReDim PrinterInfo(1 To BytesNeeded)
    
    ByteBuf = BytesNeeded
    
    'Call GetPrinter to get the status.
    Result = GetPrinter(hPrinter, 2, PrinterInfo(1), ByteBuf, _
      BytesNeeded)
    
    'Check for errors.
    If Result = 0 Then
       'Determine the error that occurred.
       LastError = err.LastDllError()
       
       'Display error message, close printer, and exit sub.
       lblCommentDesc.Caption = "Couldn't get Printer Status!  Error = " _
          & LastError
       ClosePrinter hPrinter
       Exit Sub
    End If
    
    'Copy contents of printer status byte array into a
    'PRINTER_INFO_2 structure to separate the individual elements.
    CopyMemory PI2, PrinterInfo(1), Len(PI2)
    
    'Check if printer is in ready state.
    lblStatusDesc.Caption = CheckPrinterStatus(PI2.Status)
    ' Printer info
    lblTypeDesc.Caption = GetString(PI2.pDriverName)
    lblLocationDesc.Caption = GetString(PI2.pLocation)
    lblCommentDesc.Caption = GetString(PI2.pComment)
    
    'Close the printer handle.
    ClosePrinter hPrinter
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Sub

Private Function GetString(ByVal PtrStr As Long) As String
Dim StrBuff As String * 256

SetErrModule 94, 20
If UseLocalErrorHandler Then On Error GoTo localhandler

   'Check for zero address
   If PtrStr = 0 Then
      GetString = " "
      Exit Function
   End If
   'Copy data from PtrStr to buffer.
   CopyMemory ByVal StrBuff, ByVal PtrStr, 256
   'Strip any trailing nulls from string.
   GetString = StripNulls(StrBuff)
   
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Private Function StripNulls(OriginalStr As String) As String
SetErrModule 94, 21
If UseLocalErrorHandler Then On Error GoTo localhandler

   'Strip any trailing nulls from input string.
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If
   'Return modified string.
   StripNulls = OriginalStr
   
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function

Private Function CheckPrinterStatus(PI2Status As Long) As String
Dim tempstr As String
   
SetErrModule 94, 22
If UseLocalErrorHandler Then On Error GoTo localhandler

   If PI2Status = 0 Then   ' Return "Ready"
      CheckPrinterStatus = "Ready" & vbCrLf
   Else
      tempstr = ""   ' Clear
      If (PI2Status And PRINTER_STATUS_BUSY) Then
         tempstr = tempstr & "Busy  "
      End If
      
      If (PI2Status And PRINTER_STATUS_DOOR_OPEN) Then
         tempstr = tempstr & "Printer Door Open  "
      End If
      
      If (PI2Status And PRINTER_STATUS_ERROR) Then
         tempstr = tempstr & "Printer Error  "
      End If
      
      If (PI2Status And PRINTER_STATUS_INITIALIZING) Then
         tempstr = tempstr & "Initializing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_IO_ACTIVE) Then
         tempstr = tempstr & "I/O Active  "
      End If

      If (PI2Status And PRINTER_STATUS_MANUAL_FEED) Then
         tempstr = tempstr & "Manual Feed  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NO_TONER) Then
         tempstr = tempstr & "No Toner  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NOT_AVAILABLE) Then
         tempstr = tempstr & "Not Available  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OFFLINE) Then
         tempstr = tempstr & "Off Line  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUT_OF_MEMORY) Then
         tempstr = tempstr & "Out of Memory  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempstr = tempstr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAGE_PUNT) Then
         tempstr = tempstr & "Page Punt  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_JAM) Then
         tempstr = tempstr & "Paper Jam  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_OUT) Then
         tempstr = tempstr & "Paper Out  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempstr = tempstr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_PROBLEM) Then
         tempstr = tempstr & "Page Problem  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAUSED) Then
         tempstr = tempstr & "Paused  "
      End If

      If (PI2Status And PRINTER_STATUS_PENDING_DELETION) Then
         tempstr = tempstr & "Pending Deletion  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PRINTING) Then
         tempstr = tempstr & "Printing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PROCESSING) Then
         tempstr = tempstr & "Processing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_TONER_LOW) Then
         tempstr = tempstr & "Toner Low  "
      End If

      If (PI2Status And PRINTER_STATUS_USER_INTERVENTION) Then
         tempstr = tempstr & "User Intervention  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WAITING) Then
         tempstr = tempstr & "Waiting  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WARMING_UP) Then
         tempstr = tempstr & "Warming Up  "
      End If
      
      'Did you find a known status?
      If Len(tempstr) = 0 Then
         tempstr = "Unknown Status of " & PI2Status
      End If
      
      'Return the Status
      CheckPrinterStatus = tempstr & vbCrLf
   End If
   
ResetErrModule
Exit Function

localhandler:
Dim iresponse As Integer

iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    MousePointer = vbDefault
    Exit Function
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select

End Function


