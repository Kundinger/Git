VERSION 5.00
Begin VB.Form frmTmrMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer Monitor"
   ClientHeight    =   4440
   ClientLeft      =   1650
   ClientTop       =   1650
   ClientWidth     =   10335
   Icon            =   "frmTmrMonitor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475.28
   ScaleMode       =   0  'User
   ScaleWidth      =   11535.77
   Begin VB.Frame frmTmrPerform 
      Caption         =   "System Timers"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   840
         TabIndex        =   102
         Text            =   "ph"
         Top             =   2625
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   4180
         TabIndex        =   73
         Text            =   "value"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   4880
         TabIndex        =   72
         Text            =   "actual"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   5680
         TabIndex        =   71
         Text            =   "delta"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   4185
         TabIndex        =   70
         Text            =   "value"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   4875
         TabIndex        =   69
         Text            =   "actual"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   5685
         TabIndex        =   68
         Text            =   "delta"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   4185
         TabIndex        =   67
         Text            =   "value"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   4875
         TabIndex        =   66
         Text            =   "actual"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   5685
         TabIndex        =   65
         Text            =   "delta"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   4185
         TabIndex        =   64
         Text            =   "value"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   4875
         TabIndex        =   63
         Text            =   "actual"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   5685
         TabIndex        =   62
         Text            =   "delta"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   4185
         TabIndex        =   61
         Text            =   "value"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   4875
         TabIndex        =   60
         Text            =   "actual"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   5685
         TabIndex        =   59
         Text            =   "delta"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   4185
         TabIndex        =   58
         Text            =   "value"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   4875
         TabIndex        =   57
         Text            =   "actual"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   5685
         TabIndex        =   56
         Text            =   "delta"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   4185
         TabIndex        =   55
         Text            =   "value"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   4875
         TabIndex        =   54
         Text            =   "actual"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   5685
         TabIndex        =   53
         Text            =   "delta"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   4185
         TabIndex        =   52
         Text            =   "value"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   4875
         TabIndex        =   51
         Text            =   "actual"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   5685
         TabIndex        =   50
         Text            =   "delta"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrVal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   4185
         TabIndex        =   49
         Text            =   "value"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtTmrActual 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   4875
         TabIndex        =   48
         Text            =   "actual"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtTmrDelta 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   5685
         TabIndex        =   47
         Text            =   "delta"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   6480
         TabIndex        =   46
         Text            =   "max"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   45
         Text            =   "min"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   6480
         TabIndex        =   44
         Text            =   "max"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   43
         Text            =   "min"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   42
         Text            =   "max"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   41
         Text            =   "min"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   6480
         TabIndex        =   40
         Text            =   "max"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   7200
         TabIndex        =   39
         Text            =   "min"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   6480
         TabIndex        =   38
         Text            =   "max"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   7200
         TabIndex        =   37
         Text            =   "min"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   6480
         TabIndex        =   36
         Text            =   "max"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   7200
         TabIndex        =   35
         Text            =   "min"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   6480
         TabIndex        =   34
         Text            =   "max"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   7200
         TabIndex        =   33
         Text            =   "min"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   6480
         TabIndex        =   32
         Text            =   "max"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   7200
         TabIndex        =   31
         Text            =   "min"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrMax 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   6480
         TabIndex        =   30
         Text            =   "max"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtTmrMin 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   9
         Left            =   7200
         TabIndex        =   29
         Text            =   "min"
         Top             =   2640
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   28
         Text            =   "ph"
         Top             =   720
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   27
         Text            =   "ph"
         Top             =   960
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   26
         Text            =   "ph"
         Top             =   1200
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   25
         Text            =   "ph"
         Top             =   1440
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   24
         Text            =   "ph"
         Top             =   1680
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   23
         Text            =   "ph"
         Top             =   1920
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   22
         Text            =   "ph"
         Top             =   2160
         Width           =   720
      End
      Begin VB.TextBox txtTmrPhase 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000012&
         Height          =   255
         Index           =   8
         Left            =   840
         TabIndex        =   21
         Text            =   "ph"
         Top             =   2400
         Width           =   720
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   9000
         TabIndex        =   20
         Text            =   "count"
         Top             =   720
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   8160
         TabIndex        =   19
         Text            =   "status"
         Top             =   720
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   9000
         TabIndex        =   18
         Text            =   "count"
         Top             =   960
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   2
         Left            =   8160
         TabIndex        =   17
         Text            =   "status"
         Top             =   960
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   9000
         TabIndex        =   16
         Text            =   "count"
         Top             =   1200
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   3
         Left            =   8160
         TabIndex        =   15
         Text            =   "status"
         Top             =   1200
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   9000
         TabIndex        =   14
         Text            =   "count"
         Top             =   1440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   13
         Text            =   "status"
         Top             =   1440
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   9000
         TabIndex        =   12
         Text            =   "count"
         Top             =   1680
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   5
         Left            =   8160
         TabIndex        =   11
         Text            =   "status"
         Top             =   1680
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   9000
         TabIndex        =   10
         Text            =   "count"
         Top             =   1920
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   6
         Left            =   8160
         TabIndex        =   9
         Text            =   "status"
         Top             =   1920
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   9000
         TabIndex        =   8
         Text            =   "count"
         Top             =   2160
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   7
         Left            =   8160
         TabIndex        =   7
         Text            =   "status"
         Top             =   2160
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   9000
         TabIndex        =   6
         Text            =   "count"
         Top             =   2400
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   8
         Left            =   8160
         TabIndex        =   5
         Text            =   "status"
         Top             =   2400
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox txtTmrDebugCount 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   9000
         TabIndex        =   4
         Text            =   "count"
         Top             =   2640
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox txtTmrStatus 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   9
         Left            =   8160
         TabIndex        =   3
         Text            =   "status"
         Top             =   2640
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblTmrVal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4180
         TabIndex        =   101
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblTmrDescription 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   1700
         TabIndex        =   100
         Top             =   480
         Width           =   2500
      End
      Begin VB.Label lblActual 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   4880
         TabIndex        =   99
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblTmrNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Timer#"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   480
         Width           =   640
         WordWrap        =   -1  'True
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   97
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lblDelta 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Delta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   5680
         TabIndex        =   96
         Top             =   480
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   95
         Top             =   960
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   94
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   93
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   92
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   91
         Top             =   1920
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   90
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   240
         TabIndex        =   89
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label TmrNum 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   88
         Top             =   2640
         Width           =   720
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   1
         Left            =   1700
         TabIndex        =   87
         Top             =   720
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   2
         Left            =   1700
         TabIndex        =   86
         Top             =   960
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   3
         Left            =   1700
         TabIndex        =   85
         Top             =   1200
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   4
         Left            =   1700
         TabIndex        =   84
         Top             =   1440
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   5
         Left            =   1700
         TabIndex        =   83
         Top             =   1680
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   6
         Left            =   1700
         TabIndex        =   82
         Top             =   1920
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   7
         Left            =   1700
         TabIndex        =   81
         Top             =   2160
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   8
         Left            =   1700
         TabIndex        =   80
         Top             =   2400
         Width           =   2500
      End
      Begin VB.Label lblTmrDesc 
         Caption         =   "description"
         Height          =   255
         Index           =   9
         Left            =   1700
         TabIndex        =   79
         Top             =   2640
         Width           =   2500
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   6480
         TabIndex        =   78
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   7200
         TabIndex        =   77
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblTmrPhase 
         BackStyle       =   0  'Transparent
         Caption         =   "Phase"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   940
         TabIndex        =   76
         Top             =   480
         Width           =   645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "DebugCount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   8880
         TabIndex        =   75
         Top             =   480
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   255
         Left            =   8160
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdReset 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTmrMonitor.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Reset Max & Min Values"
      Top             =   3720
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8715
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmTmrMonitor.frx":5EE4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close this screen"
      Top             =   3735
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   2640
      Top             =   3720
   End
End
Attribute VB_Name = "frmTmrMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub RefreshTimers()
Dim tmr As Integer
    For tmr = 1 To 9
        lblTmrDesc(tmr).Caption = SystemTimers(tmr).desc
        txtTmrPhase(tmr).text = Format(SystemTimers(tmr).Phase, "#0")
        txtTmrVal(tmr).text = Format(SystemTimers(tmr).Interval, "###0")
        txtTmrActual(tmr).text = Format(SystemTimers(tmr).Actual, "###0")
        txtTmrDelta(tmr).text = Format(SystemTimers(tmr).delta, "#######0")
        txtTmrMax(tmr).text = Format(SystemTimers(tmr).max, "#######0")
        txtTmrMin(tmr).text = Format(SystemTimers(tmr).Min, "#######0")
        If SystemTimers(tmr).Interval <> 0 Then
            If SystemTimers(tmr).delta < 0 Then
                txtTmrDelta(tmr).ForeColor = MEDRED
            ElseIf SystemTimers(tmr).delta >= 0 Then
                txtTmrDelta(tmr).ForeColor = Black
            End If
        End If
    Next tmr
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set frmTmrMonitor = Nothing
End Sub


Private Sub cmdReset_Click()
Dim tmr As Integer
    For tmr = 1 To 9
        SystemTimers(tmr).max = SystemTimers(tmr).delta
        SystemTimers(tmr).Min = SystemTimers(tmr).delta
    Next tmr
End Sub

Private Sub Form_Load()
Dim tmr As Integer
    ' Set Title Foreground color
    frmTmrPerform.ForeColor = Titles_ForeColor
    For tmr = 1 To 9
        If SystemTimers(tmr).Interval = 0 Then
            TmrNum(tmr).ForeColor = Common_BackColor
            lblTmrDesc(tmr).ForeColor = Common_BackColor
            txtTmrPhase(tmr).ForeColor = Common_BackColor
            txtTmrVal(tmr).ForeColor = Common_BackColor
            txtTmrActual(tmr).ForeColor = Common_BackColor
            txtTmrDelta(tmr).ForeColor = Common_BackColor
            txtTmrMax(tmr).ForeColor = Common_BackColor
            txtTmrMin(tmr).ForeColor = Common_BackColor
        End If
    Next tmr
    
    tmrUpdate.Interval = 300
End Sub

Private Sub tmrUpdate_Timer()
    RefreshTimers
End Sub
