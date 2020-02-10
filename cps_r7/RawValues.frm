VERSION 5.00
Begin VB.Form frmRawValues 
   Caption         =   "Analog RawValues"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   Icon            =   "RawValues.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7740
   ScaleWidth      =   13875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   250
      Left            =   16080
      Top             =   7680
   End
   Begin VB.Frame fraAnalogIO 
      Caption         =   "Analog I/O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   7155
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   13620
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   111
         Left            =   14880
         TabIndex        =   265
         Text            =   "7654328"
         Top             =   6480
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   110
         Left            =   14880
         TabIndex        =   264
         Text            =   "7654328"
         Top             =   6120
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   109
         Left            =   14880
         TabIndex        =   263
         Text            =   "7654328"
         Top             =   5760
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   108
         Left            =   14880
         TabIndex        =   262
         Text            =   "7654328"
         Top             =   5400
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   107
         Left            =   14880
         TabIndex        =   261
         Text            =   "7654328"
         Top             =   4920
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   106
         Left            =   14880
         TabIndex        =   260
         Text            =   "7654328"
         Top             =   4560
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   105
         Left            =   14880
         TabIndex        =   259
         Text            =   "7654328"
         Top             =   4200
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   104
         Left            =   14895
         TabIndex        =   258
         Text            =   "7654328"
         Top             =   3810
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   103
         Left            =   14880
         TabIndex        =   257
         Text            =   "7654328"
         Top             =   3360
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   102
         Left            =   14880
         TabIndex        =   256
         Text            =   "7654328"
         Top             =   3000
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   101
         Left            =   14880
         TabIndex        =   255
         Text            =   "7654328"
         Top             =   2640
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   100
         Left            =   14880
         TabIndex        =   254
         Text            =   "7654328"
         Top             =   2280
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   96
         Left            =   14880
         TabIndex        =   253
         Text            =   "7654328"
         Top             =   720
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   99
         Left            =   14880
         TabIndex        =   252
         Text            =   "7654328"
         Top             =   1800
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   98
         Left            =   14880
         TabIndex        =   251
         Text            =   "7654328"
         Top             =   1440
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   97
         Left            =   14880
         TabIndex        =   250
         Text            =   "7654328"
         Top             =   1080
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   95
         Left            =   12480
         TabIndex        =   249
         Text            =   "7654328"
         Top             =   6480
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   94
         Left            =   12480
         TabIndex        =   248
         Text            =   "7654328"
         Top             =   6120
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   93
         Left            =   12480
         TabIndex        =   247
         Text            =   "7654328"
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   92
         Left            =   12480
         TabIndex        =   246
         Text            =   "7654328"
         Top             =   5400
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   91
         Left            =   12480
         TabIndex        =   245
         Text            =   "7654328"
         Top             =   4920
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   90
         Left            =   12480
         TabIndex        =   244
         Text            =   "7654328"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   89
         Left            =   12480
         TabIndex        =   243
         Text            =   "7654328"
         Top             =   4200
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   88
         Left            =   12495
         TabIndex        =   242
         Text            =   "7654328"
         Top             =   3810
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   87
         Left            =   12480
         TabIndex        =   241
         Text            =   "7654328"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   86
         Left            =   12480
         TabIndex        =   240
         Text            =   "7654328"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   85
         Left            =   12480
         TabIndex        =   239
         Text            =   "7654328"
         Top             =   2640
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   84
         Left            =   12480
         TabIndex        =   238
         Text            =   "7654328"
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   80
         Left            =   12480
         TabIndex        =   237
         Text            =   "7654328"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   83
         Left            =   12480
         TabIndex        =   236
         Text            =   "7654328"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   82
         Left            =   12480
         TabIndex        =   235
         Text            =   "7654328"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   81
         Left            =   12480
         TabIndex        =   234
         Text            =   "7654328"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   79
         Left            =   10200
         TabIndex        =   233
         Text            =   "7654328"
         Top             =   6480
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   78
         Left            =   10200
         TabIndex        =   232
         Text            =   "7654328"
         Top             =   6120
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   77
         Left            =   10200
         TabIndex        =   231
         Text            =   "7654328"
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   76
         Left            =   10200
         TabIndex        =   230
         Text            =   "7654328"
         Top             =   5400
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   75
         Left            =   10200
         TabIndex        =   229
         Text            =   "7654328"
         Top             =   4920
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   74
         Left            =   10200
         TabIndex        =   228
         Text            =   "7654328"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   73
         Left            =   10200
         TabIndex        =   227
         Text            =   "7654328"
         Top             =   4200
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   72
         Left            =   10215
         TabIndex        =   226
         Text            =   "7654328"
         Top             =   3810
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   71
         Left            =   10200
         TabIndex        =   225
         Text            =   "7654328"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   70
         Left            =   10200
         TabIndex        =   224
         Text            =   "7654328"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   69
         Left            =   10200
         TabIndex        =   223
         Text            =   "7654328"
         Top             =   2640
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   68
         Left            =   10200
         TabIndex        =   222
         Text            =   "7654328"
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   64
         Left            =   10200
         TabIndex        =   221
         Text            =   "7654328"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   67
         Left            =   10200
         TabIndex        =   220
         Text            =   "7654328"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   66
         Left            =   10200
         TabIndex        =   219
         Text            =   "7654328"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   65
         Left            =   10200
         TabIndex        =   218
         Text            =   "7654328"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   63
         Left            =   7800
         TabIndex        =   217
         Text            =   "7654328"
         Top             =   6480
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   62
         Left            =   7800
         TabIndex        =   216
         Text            =   "7654328"
         Top             =   6120
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   61
         Left            =   7800
         TabIndex        =   215
         Text            =   "7654328"
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   60
         Left            =   7800
         TabIndex        =   214
         Text            =   "7654328"
         Top             =   5400
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   59
         Left            =   7800
         TabIndex        =   213
         Text            =   "7654328"
         Top             =   4920
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   58
         Left            =   7800
         TabIndex        =   212
         Text            =   "7654328"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   57
         Left            =   7800
         TabIndex        =   211
         Text            =   "7654328"
         Top             =   4200
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   56
         Left            =   7815
         TabIndex        =   210
         Text            =   "7654328"
         Top             =   3810
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   55
         Left            =   7800
         TabIndex        =   209
         Text            =   "7654328"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   54
         Left            =   7800
         TabIndex        =   208
         Text            =   "7654328"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   7800
         TabIndex        =   207
         Text            =   "7654328"
         Top             =   2640
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   7800
         TabIndex        =   206
         Text            =   "7654328"
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   48
         Left            =   7800
         TabIndex        =   205
         Text            =   "7654328"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   51
         Left            =   7800
         TabIndex        =   204
         Text            =   "7654328"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   50
         Left            =   7800
         TabIndex        =   203
         Text            =   "7654328"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   49
         Left            =   7800
         TabIndex        =   202
         Text            =   "7654328"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   47
         Left            =   5520
         TabIndex        =   201
         Text            =   "7654328"
         Top             =   6480
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   46
         Left            =   5520
         TabIndex        =   200
         Text            =   "7654328"
         Top             =   6120
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   45
         Left            =   5520
         TabIndex        =   199
         Text            =   "7654328"
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   5520
         TabIndex        =   198
         Text            =   "7654328"
         Top             =   5400
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   43
         Left            =   5520
         TabIndex        =   197
         Text            =   "7654328"
         Top             =   4920
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   42
         Left            =   5520
         TabIndex        =   196
         Text            =   "7654328"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   41
         Left            =   5520
         TabIndex        =   195
         Text            =   "7654328"
         Top             =   4200
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   40
         Left            =   5535
         TabIndex        =   194
         Text            =   "7654328"
         Top             =   3810
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   5520
         TabIndex        =   193
         Text            =   "7654328"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   38
         Left            =   5520
         TabIndex        =   192
         Text            =   "7654328"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   5520
         TabIndex        =   191
         Text            =   "7654328"
         Top             =   2640
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   5520
         TabIndex        =   190
         Text            =   "7654328"
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   5520
         TabIndex        =   189
         Text            =   "7654328"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   5520
         TabIndex        =   188
         Text            =   "7654328"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   5520
         TabIndex        =   187
         Text            =   "7654328"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   5520
         TabIndex        =   186
         Text            =   "7654328"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   3120
         TabIndex        =   185
         Text            =   "7654328"
         Top             =   6480
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   3120
         TabIndex        =   184
         Text            =   "7654328"
         Top             =   6120
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   3120
         TabIndex        =   183
         Text            =   "7654328"
         Top             =   5760
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   3120
         TabIndex        =   182
         Text            =   "7654328"
         Top             =   5400
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   3120
         TabIndex        =   181
         Text            =   "7654328"
         Top             =   4920
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   3120
         TabIndex        =   180
         Text            =   "7654328"
         Top             =   4560
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   3120
         TabIndex        =   179
         Text            =   "7654328"
         Top             =   4200
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   3135
         TabIndex        =   178
         Text            =   "7654328"
         Top             =   3810
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   3120
         TabIndex        =   177
         Text            =   "7654328"
         Top             =   3360
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   3120
         TabIndex        =   176
         Text            =   "7654328"
         Top             =   3000
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   3120
         TabIndex        =   175
         Text            =   "7654328"
         Top             =   2640
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   3120
         TabIndex        =   174
         Text            =   "7654328"
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3120
         TabIndex        =   173
         Text            =   "7654328"
         Top             =   720
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   3120
         TabIndex        =   172
         Text            =   "7654328"
         Top             =   1800
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   3120
         TabIndex        =   171
         Text            =   "7654328"
         Top             =   1440
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3120
         TabIndex        =   170
         Text            =   "7654328"
         Top             =   1080
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   945
         TabIndex        =   30
         Text            =   "7654328"
         Top             =   6510
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   945
         TabIndex        =   29
         Text            =   "7654328"
         Top             =   6150
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   945
         TabIndex        =   28
         Text            =   "7654328"
         Top             =   5790
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   945
         TabIndex        =   27
         Text            =   "7654328"
         Top             =   5430
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   945
         TabIndex        =   26
         Text            =   "7654328"
         Top             =   4950
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   945
         TabIndex        =   25
         Text            =   "7654328"
         Top             =   4590
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   945
         TabIndex        =   24
         Text            =   "7654328"
         Top             =   4230
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   960
         TabIndex        =   23
         Text            =   "7654328"
         Top             =   3840
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   945
         TabIndex        =   22
         Text            =   "7654328"
         Top             =   3390
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   945
         TabIndex        =   21
         Text            =   "7654328"
         Top             =   3030
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   945
         TabIndex        =   20
         Text            =   "7654328"
         Top             =   2670
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   945
         TabIndex        =   19
         Text            =   "7654328"
         Top             =   2310
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   945
         TabIndex        =   18
         Text            =   "7654328"
         Top             =   750
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   945
         TabIndex        =   17
         Text            =   "7654328"
         Top             =   1830
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   945
         TabIndex        =   16
         Text            =   "7654328"
         Top             =   1470
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   945
         TabIndex        =   15
         Text            =   "7654328"
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   112
         Left            =   17085
         TabIndex        =   14
         Text            =   "7654328"
         Top             =   3390
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   113
         Left            =   17085
         TabIndex        =   13
         Text            =   "7654328"
         Top             =   3030
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   114
         Left            =   17085
         TabIndex        =   12
         Text            =   "7654328"
         Top             =   2670
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   115
         Left            =   17085
         TabIndex        =   11
         Text            =   "7654328"
         Top             =   2310
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   116
         Left            =   17085
         TabIndex        =   10
         Text            =   "7654328"
         Top             =   1830
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   117
         Left            =   17085
         TabIndex        =   9
         Text            =   "7654328"
         Top             =   1470
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   118
         Left            =   17085
         TabIndex        =   8
         Text            =   "7654328"
         Top             =   1110
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   119
         Left            =   17085
         TabIndex        =   7
         Text            =   "7654328"
         Top             =   750
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   124
         Left            =   17085
         TabIndex        =   6
         Text            =   "7654328"
         Top             =   3870
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   125
         Left            =   17085
         TabIndex        =   5
         Text            =   "7654328"
         Top             =   4950
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   126
         Left            =   17085
         TabIndex        =   4
         Text            =   "7654328"
         Top             =   4590
         Width           =   900
      End
      Begin VB.TextBox txtReadRaw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   127
         Left            =   17085
         TabIndex        =   3
         Text            =   "7654328"
         Top             =   4230
         Width           =   900
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   169
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   108
         Left            =   14040
         TabIndex        =   168
         Top             =   5430
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   109
         Left            =   14040
         TabIndex        =   167
         Top             =   5790
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   110
         Left            =   14040
         TabIndex        =   166
         Top             =   6150
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   111
         Left            =   14040
         TabIndex        =   165
         Top             =   6510
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   14
         Left            =   14040
         TabIndex        =   164
         Top             =   360
         Visible         =   0   'False
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   105
         Left            =   14040
         TabIndex        =   163
         Top             =   4260
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   106
         Left            =   14040
         TabIndex        =   162
         Top             =   4620
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   107
         Left            =   14040
         TabIndex        =   161
         Top             =   4980
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   104
         Left            =   14040
         TabIndex        =   160
         Top             =   3900
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   14
         Left            =   14820
         TabIndex        =   159
         Top             =   480
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   96
         Left            =   14040
         TabIndex        =   158
         Top             =   780
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   97
         Left            =   14040
         TabIndex        =   157
         Top             =   1140
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   98
         Left            =   14040
         TabIndex        =   156
         Top             =   1500
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   99
         Left            =   14040
         TabIndex        =   155
         Top             =   1860
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   100
         Left            =   14040
         TabIndex        =   154
         Top             =   2340
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   101
         Left            =   14040
         TabIndex        =   153
         Top             =   2700
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   102
         Left            =   14040
         TabIndex        =   152
         Top             =   3060
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   103
         Left            =   14040
         TabIndex        =   151
         Top             =   3420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   11
         Left            =   11760
         TabIndex        =   150
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   89
         Left            =   11760
         TabIndex        =   149
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   90
         Left            =   11760
         TabIndex        =   148
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   91
         Left            =   11760
         TabIndex        =   147
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   88
         Left            =   11760
         TabIndex        =   146
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   11
         Left            =   12540
         TabIndex        =   145
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   92
         Left            =   11760
         TabIndex        =   144
         Top             =   5460
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   93
         Left            =   11760
         TabIndex        =   143
         Top             =   5820
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   94
         Left            =   11760
         TabIndex        =   142
         Top             =   6180
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   95
         Left            =   11760
         TabIndex        =   141
         Top             =   6540
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   80
         Left            =   11760
         TabIndex        =   140
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   81
         Left            =   11760
         TabIndex        =   139
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   82
         Left            =   11760
         TabIndex        =   138
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   83
         Left            =   11760
         TabIndex        =   137
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   84
         Left            =   11760
         TabIndex        =   136
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   85
         Left            =   11760
         TabIndex        =   135
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   86
         Left            =   11760
         TabIndex        =   134
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   87
         Left            =   11760
         TabIndex        =   133
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Left            =   180
         TabIndex        =   132
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Left            =   180
         TabIndex        =   131
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Left            =   180
         TabIndex        =   130
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Left            =   180
         TabIndex        =   129
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Left            =   180
         TabIndex        =   128
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Left            =   180
         TabIndex        =   127
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Left            =   180
         TabIndex        =   126
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   0
         Left            =   180
         TabIndex        =   125
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   15
         Left            =   180
         TabIndex        =   124
         Top             =   6540
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   14
         Left            =   180
         TabIndex        =   123
         Top             =   6180
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   13
         Left            =   180
         TabIndex        =   122
         Top             =   5820
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   12
         Left            =   180
         TabIndex        =   121
         Top             =   5460
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   2
         Left            =   960
         TabIndex        =   120
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Left            =   180
         TabIndex        =   119
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   11
         Left            =   180
         TabIndex        =   118
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   10
         Left            =   180
         TabIndex        =   117
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Left            =   180
         TabIndex        =   116
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   23
         Left            =   2400
         TabIndex        =   115
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   22
         Left            =   2400
         TabIndex        =   114
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   21
         Left            =   2400
         TabIndex        =   113
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   20
         Left            =   2400
         TabIndex        =   112
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   19
         Left            =   2400
         TabIndex        =   111
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   18
         Left            =   2400
         TabIndex        =   110
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   17
         Left            =   2400
         TabIndex        =   109
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   16
         Left            =   2400
         TabIndex        =   108
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   31
         Left            =   2400
         TabIndex        =   107
         Top             =   6540
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   30
         Left            =   2400
         TabIndex        =   106
         Top             =   6180
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   29
         Left            =   2400
         TabIndex        =   105
         Top             =   5820
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   28
         Left            =   2400
         TabIndex        =   104
         Top             =   5460
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   3
         Left            =   3180
         TabIndex        =   103
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   24
         Left            =   2400
         TabIndex        =   102
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   27
         Left            =   2400
         TabIndex        =   101
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   26
         Left            =   2400
         TabIndex        =   100
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   25
         Left            =   2400
         TabIndex        =   99
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   3
         Left            =   2400
         TabIndex        =   98
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   39
         Left            =   4800
         TabIndex        =   97
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   38
         Left            =   4800
         TabIndex        =   96
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   37
         Left            =   4800
         TabIndex        =   95
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   36
         Left            =   4800
         TabIndex        =   94
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   35
         Left            =   4800
         TabIndex        =   93
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   34
         Left            =   4800
         TabIndex        =   92
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   33
         Left            =   4800
         TabIndex        =   91
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   32
         Left            =   4800
         TabIndex        =   90
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   6
         Left            =   5580
         TabIndex        =   89
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   40
         Left            =   4800
         TabIndex        =   88
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   43
         Left            =   4800
         TabIndex        =   87
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   42
         Left            =   4800
         TabIndex        =   86
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   41
         Left            =   4800
         TabIndex        =   85
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   84
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   55
         Left            =   7080
         TabIndex        =   83
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   54
         Left            =   7080
         TabIndex        =   82
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   53
         Left            =   7080
         TabIndex        =   81
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   52
         Left            =   7080
         TabIndex        =   80
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   51
         Left            =   7080
         TabIndex        =   79
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   50
         Left            =   7080
         TabIndex        =   78
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   49
         Left            =   7080
         TabIndex        =   77
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   48
         Left            =   7080
         TabIndex        =   76
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   7
         Left            =   7860
         TabIndex        =   75
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   56
         Left            =   7080
         TabIndex        =   74
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   59
         Left            =   7080
         TabIndex        =   73
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   58
         Left            =   7080
         TabIndex        =   72
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   57
         Left            =   7080
         TabIndex        =   71
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   5
         Left            =   7080
         TabIndex        =   70
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   71
         Left            =   9480
         TabIndex        =   69
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   70
         Left            =   9480
         TabIndex        =   68
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   69
         Left            =   9480
         TabIndex        =   67
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   68
         Left            =   9480
         TabIndex        =   66
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   67
         Left            =   9480
         TabIndex        =   65
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   66
         Left            =   9480
         TabIndex        =   64
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   65
         Left            =   9480
         TabIndex        =   63
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   64
         Left            =   9480
         TabIndex        =   62
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblCounts 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Counts"
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
         Index           =   10
         Left            =   10260
         TabIndex        =   61
         Top             =   480
         Width           =   795
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   72
         Left            =   9480
         TabIndex        =   60
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   75
         Left            =   9480
         TabIndex        =   59
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   74
         Left            =   9480
         TabIndex        =   58
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   73
         Left            =   9480
         TabIndex        =   57
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   10
         Left            =   9480
         TabIndex        =   56
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   112
         Left            =   16320
         TabIndex        =   55
         Top             =   3420
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   113
         Left            =   16320
         TabIndex        =   54
         Top             =   3060
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   114
         Left            =   16320
         TabIndex        =   53
         Top             =   2700
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   115
         Left            =   16320
         TabIndex        =   52
         Top             =   2340
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   116
         Left            =   16320
         TabIndex        =   51
         Top             =   1860
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   117
         Left            =   16320
         TabIndex        =   50
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   118
         Left            =   16320
         TabIndex        =   49
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   119
         Left            =   16320
         TabIndex        =   48
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "48/12"
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
         Index           =   124
         Left            =   16320
         TabIndex        =   47
         Top             =   3900
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   125
         Left            =   16320
         TabIndex        =   46
         Top             =   4980
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   126
         Left            =   16320
         TabIndex        =   45
         Top             =   4620
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   127
         Left            =   16320
         TabIndex        =   44
         Top             =   4260
         Width           =   585
      End
      Begin VB.Label lblAdChn 
         BackStyle       =   0  'Transparent
         Caption         =   " Addr /Chan"
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
         Height          =   375
         Index           =   7
         Left            =   16320
         TabIndex        =   43
         Top             =   360
         Width           =   585
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   47
         Left            =   4800
         TabIndex        =   42
         Top             =   6510
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   46
         Left            =   4800
         TabIndex        =   41
         Top             =   6150
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   45
         Left            =   4800
         TabIndex        =   40
         Top             =   5790
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   44
         Left            =   4800
         TabIndex        =   39
         Top             =   5430
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   63
         Left            =   7080
         TabIndex        =   38
         Top             =   6510
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   62
         Left            =   7080
         TabIndex        =   37
         Top             =   6150
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   61
         Left            =   7080
         TabIndex        =   36
         Top             =   5790
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   60
         Left            =   7080
         TabIndex        =   35
         Top             =   5430
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   79
         Left            =   9480
         TabIndex        =   34
         Top             =   6510
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   78
         Left            =   9480
         TabIndex        =   33
         Top             =   6150
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   77
         Left            =   9480
         TabIndex        =   32
         Top             =   5790
         Width           =   585
      End
      Begin VB.Label lblAddrChan 
         BackStyle       =   0  'Transparent
         Caption         =   "2/12"
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
         Index           =   76
         Left            =   9480
         TabIndex        =   31
         Top             =   5430
         Width           =   585
      End
   End
   Begin VB.TextBox txtReadRaw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   120
      Left            =   15405
      TabIndex        =   0
      Text            =   "7654328"
      Top             =   13590
      Width           =   900
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Raw Values from I/O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   4043
      TabIndex        =   266
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblAddrChan 
      BackStyle       =   0  'Transparent
      Caption         =   "2/12"
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
      Index           =   120
      Left            =   14640
      TabIndex        =   1
      Top             =   13620
      Width           =   585
   End
End
Attribute VB_Name = "frmRawValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrUpdate_Timer()
Dim addr As Integer
Dim chan As Integer
Dim iAddr As Integer
Dim iChan As Integer
Dim iStn As Integer
Dim iOffset As Integer
Dim idx As Integer
For iStn = 0 To LAST_STN
    For iOffset = 2 To 3
        iAddr = (iStn * 4) + iOffset
        For iChan = 0 To 15
            idx = (((iStn) * 32) + ((iOffset - 2) * 16)) + iChan
            lblAddrChan(idx).Caption = Format(iAddr, "##0") & "/" & Format(iChan, "00")
            txtReadRaw(idx).text = Format(OptoAIO(iAddr, iChan).RawValue, "###,###,##0")
        Next iChan
    Next iOffset
Next iStn
End Sub
