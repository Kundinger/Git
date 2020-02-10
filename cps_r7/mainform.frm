VERSION 5.00
Begin VB.Form frmMainForm 
   Caption         =   "Mistic I/O User's Utility"
   ClientHeight    =   8640
   ClientLeft      =   1725
   ClientTop       =   3285
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8640
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRawValues 
      Caption         =   "Opto Analog Raw Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   94
      ToolTipText     =   "Open Opto22 Error Monitor Screen"
      Top             =   7920
      Width           =   7080
   End
   Begin VB.CommandButton cmdCommOff 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2663
      MaskColor       =   &H00FFFFFF&
      Picture         =   "mainform.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   6300
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton cmdCommOn 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3651
      MaskColor       =   &H00FFFFFF&
      Picture         =   "mainform.frx":1B42
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   6300
      UseMaskColor    =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton cmdOptoErrors 
      Caption         =   "Opto Error Monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   91
      ToolTipText     =   "Open Opto22 Error Monitor Screen"
      Top             =   7200
      Width           =   7080
   End
   Begin VB.TextBox txterrorscount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   87
      Text            =   "0"
      Top             =   6540
      Width           =   2455
   End
   Begin VB.TextBox txtCount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4640
      TabIndex        =   85
      Text            =   "0"
      Top             =   6540
      Width           =   2455
   End
   Begin VB.Frame ParamFrame 
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdOPTOConfiguration 
         Caption         =   "&OPTO Brain Configure Button"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   90
         Top             =   5520
         Width           =   4695
      End
      Begin VB.CommandButton cmdPowerUpClear 
         Caption         =   "Power-up clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   84
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton btnHelp 
         Caption         =   "&Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton btnConfigPort 
         Caption         =   "Configure &Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   82
         Top             =   2640
         Width           =   1572
      End
      Begin VB.Frame PosFrame 
         Caption         =   "Positions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1212
         Left            =   120
         TabIndex        =   79
         Top             =   1320
         Width           =   1572
         Begin VB.TextBox editPos 
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   8
            TabIndex        =   3
            Text            =   "0"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox editPos 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   8
            TabIndex        =   2
            Text            =   "0"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            Caption         =   "( 1 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   81
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label3 
            Caption         =   "( 0 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   372
         End
      End
      Begin VB.Frame Frame2 
         Height          =   972
         Left            =   120
         TabIndex        =   76
         Top             =   240
         Width           =   1572
         Begin VB.TextBox editCommand 
            Height          =   288
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   1
            Text            =   "0"
            Top             =   600
            Width           =   372
         End
         Begin VB.TextBox editAddress 
            Height          =   288
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "0"
            Top             =   240
            Width           =   372
         End
         Begin VB.Label Label2 
            Caption         =   "Command"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   77
            Top             =   240
            Width           =   732
         End
      End
      Begin VB.Frame SendFrame 
         Caption         =   "Send Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4215
         Left            =   1920
         TabIndex        =   58
         Top             =   240
         Width           =   2175
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   15
            Left            =   600
            MaxLength       =   15
            TabIndex        =   19
            Text            =   "0"
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   14
            Left            =   600
            MaxLength       =   15
            TabIndex        =   18
            Text            =   "0"
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   13
            Left            =   600
            MaxLength       =   15
            TabIndex        =   17
            Text            =   "0"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   12
            Left            =   600
            MaxLength       =   15
            TabIndex        =   16
            Text            =   "0"
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   11
            Left            =   600
            MaxLength       =   15
            TabIndex        =   15
            Text            =   "0"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   10
            Left            =   600
            MaxLength       =   15
            TabIndex        =   14
            Text            =   "0"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   9
            Left            =   600
            MaxLength       =   15
            TabIndex        =   13
            Text            =   "0"
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   8
            Left            =   600
            MaxLength       =   15
            TabIndex        =   12
            Text            =   "0"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   7
            Left            =   600
            MaxLength       =   15
            TabIndex        =   11
            Text            =   "0"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   6
            Left            =   600
            MaxLength       =   15
            TabIndex        =   10
            Text            =   "0"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   5
            Left            =   600
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "0"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   4
            Left            =   600
            MaxLength       =   15
            TabIndex        =   8
            Text            =   "0"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   3
            Left            =   600
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "0"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   2
            Left            =   600
            MaxLength       =   15
            TabIndex        =   6
            Text            =   "0"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   15
            TabIndex        =   5
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox editSend 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   15
            TabIndex        =   4
            Text            =   "0"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "( 0 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   75
            Top             =   240
            Width           =   372
         End
         Begin VB.Label Label6 
            Caption         =   "( 1 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   74
            Top             =   480
            Width           =   372
         End
         Begin VB.Label Label8 
            Caption         =   "( 2 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   73
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label9 
            Caption         =   "( 3 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   72
            Top             =   960
            Width           =   372
         End
         Begin VB.Label Label7 
            Caption         =   "( 4 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   71
            Top             =   1200
            Width           =   372
         End
         Begin VB.Label Label10 
            Caption         =   "( 5 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   70
            Top             =   1440
            Width           =   372
         End
         Begin VB.Label Label11 
            Caption         =   "( 6 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   69
            Top             =   1680
            Width           =   372
         End
         Begin VB.Label Label12 
            Caption         =   "( 7 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   68
            Top             =   1920
            Width           =   372
         End
         Begin VB.Label Label13 
            Caption         =   "( 8 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   67
            Top             =   2160
            Width           =   372
         End
         Begin VB.Label Label14 
            Caption         =   "( 9 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   66
            Top             =   2400
            Width           =   372
         End
         Begin VB.Label Label15 
            Caption         =   "( 10 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   65
            Top             =   2640
            Width           =   492
         End
         Begin VB.Label Label16 
            Caption         =   "( 11 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   64
            Top             =   2880
            Width           =   492
         End
         Begin VB.Label Label17 
            Caption         =   "( 12 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   63
            Top             =   3120
            Width           =   492
         End
         Begin VB.Label Label18 
            Caption         =   "( 13 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   62
            Top             =   3360
            Width           =   492
         End
         Begin VB.Label Label19 
            Caption         =   "( 14 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   61
            Top             =   3600
            Width           =   492
         End
         Begin VB.Label Label20 
            Caption         =   "( 15 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   60
            Top             =   3840
            Width           =   492
         End
      End
      Begin VB.Frame ReceFrame 
         Caption         =   "ReceData"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   4320
         TabIndex        =   39
         Top             =   240
         Width           =   2295
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   15
            Left            =   600
            MaxLength       =   15
            TabIndex        =   35
            Text            =   "0"
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   14
            Left            =   600
            MaxLength       =   15
            TabIndex        =   34
            Text            =   "0"
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   13
            Left            =   600
            MaxLength       =   15
            TabIndex        =   33
            Text            =   "0"
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   12
            Left            =   600
            MaxLength       =   15
            TabIndex        =   32
            Text            =   "0"
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   11
            Left            =   600
            MaxLength       =   15
            TabIndex        =   31
            Text            =   "0"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   10
            Left            =   600
            MaxLength       =   15
            TabIndex        =   30
            Text            =   "0"
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   9
            Left            =   600
            MaxLength       =   15
            TabIndex        =   29
            Text            =   "0"
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   8
            Left            =   600
            MaxLength       =   15
            TabIndex        =   28
            Text            =   "0"
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   7
            Left            =   600
            MaxLength       =   15
            TabIndex        =   27
            Text            =   "0"
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   6
            Left            =   600
            MaxLength       =   15
            TabIndex        =   26
            Text            =   "0"
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   5
            Left            =   600
            MaxLength       =   15
            TabIndex        =   25
            Text            =   "0"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   4
            Left            =   600
            MaxLength       =   15
            TabIndex        =   24
            Text            =   "0"
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   3
            Left            =   600
            MaxLength       =   15
            TabIndex        =   23
            Text            =   "0"
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   2
            Left            =   600
            MaxLength       =   15
            TabIndex        =   22
            Text            =   "0"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   15
            TabIndex        =   21
            Text            =   "0"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox editRece 
            Height          =   285
            Index           =   0
            Left            =   600
            MaxLength       =   15
            TabIndex        =   20
            Text            =   "0"
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label21 
            Caption         =   "( 15 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   57
            Top             =   3840
            Width           =   492
         End
         Begin VB.Label Label22 
            Caption         =   "( 14 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   55
            Top             =   3600
            Width           =   492
         End
         Begin VB.Label Label23 
            Caption         =   "( 13 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   53
            Top             =   3360
            Width           =   492
         End
         Begin VB.Label Label24 
            Caption         =   "( 12 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   52
            Top             =   3120
            Width           =   492
         End
         Begin VB.Label Label25 
            Caption         =   "( 11 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   51
            Top             =   2880
            Width           =   492
         End
         Begin VB.Label Label26 
            Caption         =   "( 10 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   50
            Top             =   2640
            Width           =   492
         End
         Begin VB.Label Label27 
            Caption         =   "( 9 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   49
            Top             =   2400
            Width           =   372
         End
         Begin VB.Label Label28 
            Caption         =   "( 8 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   48
            Top             =   2160
            Width           =   372
         End
         Begin VB.Label Label29 
            Caption         =   "( 7 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   372
         End
         Begin VB.Label Label30 
            Caption         =   "( 6 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   46
            Top             =   1680
            Width           =   372
         End
         Begin VB.Label Label31 
            Caption         =   "( 5 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   45
            Top             =   1440
            Width           =   372
         End
         Begin VB.Label Label32 
            Caption         =   "( 4 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   372
         End
         Begin VB.Label Label33 
            Caption         =   "( 3 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   43
            Top             =   960
            Width           =   372
         End
         Begin VB.Label Label34 
            Caption         =   "( 2 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   372
         End
         Begin VB.Label Label35 
            Caption         =   "( 1 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   41
            Top             =   480
            Width           =   372
         End
         Begin VB.Label Label36 
            Caption         =   "( 0 )"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   372
         End
      End
      Begin VB.CommandButton btnSend 
         Caption         =   "&Send"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   54
         Top             =   3120
         Width           =   1572
      End
      Begin VB.CommandButton btnExit 
         Caption         =   "&Return"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   59
         Top             =   5280
         Width           =   1572
      End
      Begin VB.CommandButton btnClear 
         Caption         =   "&Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         TabIndex        =   56
         Top             =   3600
         Width           =   1572
      End
      Begin VB.Frame ErrorFrame 
         Caption         =   "Error Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1860
         TabIndex        =   37
         Top             =   4620
         Width           =   4695
         Begin VB.Label ErrorLbl 
            Caption         =   "None"
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Width           =   4455
         End
      End
   End
   Begin VB.Label Label38 
      Caption         =   "Good Messages"
      Height          =   255
      Left            =   4640
      TabIndex        =   89
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Label Label37 
      Caption         =   "Errors  Messages"
      Height          =   255
      Left            =   120
      TabIndex        =   88
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label txtErrorCount 
      Caption         =   "None"
      Height          =   315
      Left            =   2160
      TabIndex        =   86
      Top             =   5340
      Width           =   4455
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 93 '''''''''' Form MAINFORM.frm ''''''''''''''''
Option Explicit

Private Sub btnClear_Click()
' declare local variable
Dim i As Single
    ' reset all values to 0
    editAddress.text = "0"
    editCommand.text = "0"
    editPos(0).text = "0"
    editPos(1).text = "0"
    For i = 0 To 15
        editSend(i).text = "0"
        editRece(i).text = "0"
    Next i
    ErrorLbl.Caption = "None"
End Sub

Private Sub btnConfigPort_Click()
    PortConfigForm.Show 1
End Sub

Private Sub btnExit_Click()
    Visible = False
    frmMainMenu.Show
End Sub

Private Sub btnHelp_Click()
   ' ViewFile = FILEPATH + "mistuser_help.txt"
   ' Load frmHelpForm
   ' frmHelpForm.Show 1
End Sub

Private Sub btnSend_Click()              'Manual file I/O's
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 1
Start:
    Brick_Address = Val(editAddress)
    Brick_Cmd = Val(editCommand)
    Brick_Pos(0) = Val(editPos(0))
    Brick_Pos(1) = Val(editPos(1))
    
    Dim i As Single
    For i = 0 To 15
        Brick_Send(i) = Val(editSend(i))
    Next i                              ' send the command
    Brick_Error = SendMIO(Brick_Handle, _
                          Brick_Address, _
                          Brick_Cmd, _
                          Brick_Pos(0), _
                          Brick_Send(0), _
                          Brick_Rece(0))
                        
    Dim ErrTxt$
    If Brick_Error <> "0" Then                     ' All errors are negetive
      txterrorscount = txterrorscount + 1         ' Increment the errors counter
    Else
      ErrorLbl.Caption = "none"
      txtCount = txtCount + 1                       ' Increment the message count
    End If
    ErrTxt$ = MisticError(Brick_Error)
    ErrorLbl.Caption = str$(Brick_Error) + "   " + ErrTxt$
    DoEvents                                     ' Let others have the processor
    For i = 0 To 15
        editRece(i).text = str$(Brick_Rece(i))  ' display the return values
    Next i
    ' ***************** TESTING CODE TYPE *************************
 '   Dim count As Long
 '   If Brick_Address = 3 Then
 '    If Str$(Brick_Rece(4)) < 160000 Or Str$(Brick_Rece(5)) < 160000 Then
 '     editAddress = 7 ' WAS 7
 '     txterrorscount = txterrorscount + 1         ' Increment the errors counter
 '     Else
 '     editAddress = 7 ' WAS 7
 '    End If
 '   Else
 '   If Brick_Address = 15 Then
 '    If Str$(Brick_Rece(2)) > -1000 Or Str$(Brick_Rece(3)) > -3500 Then
 '     txterrorscount = txterrorscount + 1         ' Increment the errors counter
 '     editAddress = 3  'WAS 3
 '     Else
 '     editAddress = 3  'WAS 3
 '    End If
 '   Else
 '   If Brick_Address = 7 Then
 '    If Str$(Brick_Rece(2)) > -0 Or Str$(Brick_Rece(3)) < 1500 Then
 '     txterrorscount = txterrorscount + 1         ' Increment the errors counter
 '     editAddress = 15  'WAS 15
 '     Else
 '     editAddress = 15  'WAS 15
 '    End If
 '   End If
 '   End If
 '   End If
 '     For count = 0 To 500000
 '       count = count + 1
 '     Next count
 '   '  DoEvents
 ' '  GoTo Start
 ' ****************** TEST OUT *****************************
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

Private Sub cmdCommOff_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 139
    If (IoComOn) Then
        ' Turn Off Opto22 I/O Communications
        OptoCommOff_Request = True
    End If
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

Private Sub cmdCommOn_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 138
    If (Not IoComOn) Then
        ' Start Opto22 I/O Communications
        OptoCommOn_Request = True
    End If
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

Private Sub cmdOPTOConfiguration_Click()
'
'   Configure the brain fix
'
'   When no errors exist save all parms to the brain
'       See OPTO book page 26
'

Dim count As Integer
Dim brainnumber As Integer
Dim i As Integer
Dim ErrTxt$

    brainnumber = 0
    If txterrorscount = VALUE0 Then
        For count = 0 To NR_STN
            If Node_Info(count) > 0 Then
                If count > 0 Then
                    brainnumber = count * 4
                End If
                Brick_Address = Val(brainnumber)
                Brick_Cmd = Val(3)
                Brick_Pos(0) = Val(0)
                Brick_Pos(1) = Val(0)
                
                For i = 0 To 15
                    Brick_Send(i) = Val(0)
                Next i                                              ' send the command
                Brick_Error = SendMIO(Brick_Handle, _
                                  Brick_Address, _
                                  Brick_Cmd, _
                                  Brick_Pos(0), _
                                  Brick_Send(0), _
                                  Brick_Rece(0))
                                
                If Brick_Error <> "0" Then                          ' All errors are negative
                    txterrorscount = txterrorscount + 1             ' Increment the errors counter
                Else
                    ErrorLbl.Caption = "none"
                    txtCount = txtCount + 1                         ' Increment the message count
                End If
                ErrTxt$ = MisticError(Brick_Error)
                ErrorLbl.Caption = str$(Brick_Error) + "   " + ErrTxt$
                DoEvents                                            ' Let others have the processor
            End If
        Next count
        Delay_Box "Configurations Saved.", MSGDELAY, msgSHOW
    Else
        Delay_Box "Can Not save configuration while errors exist.", MSGDELAY, msgSHOW
    End If

End Sub

Private Sub cmdOptoErrors_Click()
'  Unload Me
  frmOptoErrors.Show
End Sub

Private Sub cmdPowerUpClear_Click()
    PowerUPClear
End Sub

Private Sub cmdRawValues_Click()
'    frmRawValues.Show
End Sub

Private Sub Form_Load()
    If IoComOn Then
        If btnSend.Enabled <> True Then       ' Setup I/O's and clear everything
            ' Only do this one time at power up
            ' App.HelpFile = "c:\cps_r7\Misticware.hlp"
            Load PortConfigForm
            PortConfigForm.btnPortOk = True
            Unload PortConfigForm
            Visible = False
        Else
            Visible = True
        End If
        ' update command buttons
        cmdCommOn.Enabled = False
        cmdCommOff.Enabled = True
    Else
        ' update command buttons
        cmdCommOn.Enabled = True
        cmdCommOff.Enabled = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmMainForm = Nothing    'current form
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Public Sub OptoCommOff()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 1390
    If (IoComOn) Then
        ' Reset Watchdogs for All Boards to 1 sec
        WatchdogSetup 1
        ' turn Off Opto22 I/O Communications
        IoComOn = False
        ' update command buttons
        cmdCommOn.Enabled = True
        cmdCommOff.Enabled = False
        OptoCommOff_Request = False
    End If
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

Public Sub OptoCommOn()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 1380
    If (Not IoComOn) Then
        ' Starting Opto22 I/O Communications
        IoComOn = True
        ' Restart All Brains
        PowerUPClear
        ' update command buttons
        cmdCommOn.Enabled = False
        cmdCommOff.Enabled = True
        OptoCommOn_Request = False
    End If
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

Function Send_Opto_Command(Opto_Address, Opto_Command, Opto_Pos0, Opto_Pos1)

Dim iii As Single              'Will fill commands to the OPTO Brick
Dim ii As Single               'putting error in Opto_Error_data
Dim i As Single                'putting receive data in Opto_Rec_data
Dim ErrTxt$
Dim OneTimeError As Integer
Dim goodstuff As Integer
Dim boxFlag As Boolean

If UseLocalErrorHandler Then On Error GoTo localhandler
If Not IoComOn Then Exit Function
SetErrModule 93, 2




'            ChgErrModule 93, 2210
    editAddress = Opto_Address
'            ChgErrModule 93, 2211
    editCommand = Opto_Command
'            ChgErrModule 93, 2212
    editPos(0) = Opto_Pos0
'            ChgErrModule 93, 2213
    editPos(1) = Opto_Pos1

'            ChgErrModule 93, 2218
    For iii = 0 To 15
        editSend(iii).text = Opto_Send_Data(iii)
    Next iii
'            ChgErrModule 93, 2219
    For i = 0 To 15
           Brick_Send(i) = Val(editSend(i))
    Next i
    OneTimeError = 0
 
'            ChgErrModule 93, 2220
            
OneTimeErrorLoop:
 
            ChgErrModule 93, 2221
            
    Brick_Address = Val(editAddress)
    Brick_Cmd = Val(editCommand)
    Brick_Pos(0) = Val(editPos(0))
    Brick_Pos(1) = Val(editPos(1))
    Brick_Error = SendMIO(Brick_Handle, _
                          Brick_Address, _
                          Brick_Cmd, _
                          Brick_Pos(0), _
                          Brick_Send(0), _
                          Brick_Rece(0))
    ErrTxt$ = MisticError(Brick_Error)
 
            ChgErrModule 93, 2222
            
    If Brick_Error <> "0" Then                          ' All errors are negative
       If OneTimeError < 3 Then                         ' See if it's erronius could do 9 times
          OneTimeError = OneTimeError + 1
          GoTo OneTimeErrorLoop
       Else                                             ' Tried three times and still an error
          txterrorscount = txterrorscount + 1           ' Increment the errors counter
          boxFlag = IIf(((txterrorscount Mod 128) = 0), True, False)
          ErrTxt$ = MisticError(Brick_Error)            ' Display them for now
          ErrorLbl.Caption = str$(Brick_Error) + "   " + ErrTxt$
          Opto_Error_data = str$(Brick_Error) + "   " + ErrTxt$
          ErrTxt$ = ErrTxt$ _
                        + "(cmd=" + editCommand _
                        + " addr=" + CStr(editAddress) _
                        + " chan=" + CStr(editPos(0)) _
                        + " sent=" + CStr(editSend(0)) _
                        + " recd=" + CStr(Brick_Rece(0)) _
                        + ")"
          If boxFlag Then Delay_Box ErrTxt$, MSGDELAY, msgSHOW          '   + " opto error"
          Write_ELog ErrTxt$                            '   + " opto error"
          
                ChgErrModule 93, 2223
                
          Opto_Rec_Data(16) = ErrTxt$                   '  error flag in bit 16
          IoComm_Flag = False
       End If
    Else
         ErrorLbl.Caption = "none"
         Opto_Rec_Data(16) = " "
    
         txtCount = txtCount + 1                        ' Increment the message count
         goodstuff = 0
       
         For i = 0 To 15
           editRece(i).text = str$(Brick_Rece(i))
           Opto_Rec_Data(i) = str$(Brick_Rece(i))
           If str$(Brick_Rece(i)) > 0 Then
               goodstuff = 1
           End If
         Next i
    End If
      
            ChgErrModule 93, 2224
            
    If goodstuff = 0 Then
        Opto_Rec_Data(16) = " "
    End If
 
 
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

Function WatchdogSetup(ByVal secDelay As Long)
' Command 114=Set Watchdog
'
'
Dim i As Integer
Dim ii As Integer
Dim base As Integer
Dim addr As Integer
Dim chan As Integer
Dim wdDelay As Double
Dim pos0 As Long
Dim pos1 As Long

    ' Reset all the digital Watchdogs after loss of comm
'    Opto_Send_Data(0) = val(3200)                          ' set Watchdog Time to 32 sec (3200 * 10msec)
    wdDelay = CDbl(100 * secDelay)
    Opto_Send_Data(0) = wdDelay                             ' set Watchdog Time to n sec (n00 * 10msec)
    
    ' Common Board
    pos0 = 0
    pos1 = 0
    For i = 0 To 3
        For ii = 0 To 3
            chan = ((4 * i) + ii)
                If (chan <> Com_DIO(icAlarmBeacon).chan) Then
                    pos1 = pos1 + 2 ^ chan
                Else
                    pos0 = 2 ^ chan
                End If
        Next ii
    Next i
    
    ' Main Board Modules
'    Send_Opto_Command 0, 114, 8192, 57343                   ' main module 0 (All Off except Beacon; Turn Beacon On)
    Send_Opto_Command 0, 114, pos0, pos1                    ' main module 0 (All Off except Beacon; Turn Beacon On)
    Send_Opto_Command 1, 114, 0, OptoChanMask(1)            ' main module 1 (Some Off)
    
    ' Station Modules
    For i = 1 To NR_STN
        If Node_Info(i) > 0 Then
            base = 4 * i
            For ii = 0 To 1
                addr = base + ii                            ' station modules 0 & 1 (Some Off)
                If OptoChanMask(addr) > 0 Then Send_Opto_Command addr, 114, 0, OptoChanMask(addr)
            Next ii
        End If
    Next i
    
    ' Clear SendData
    Opto_Send_Data(0) = Val(0)

End Function

Function PowerUPClear()

    ' Order is Important / General Clear First, Initialize Digitals second, then Initialize Analogs
    If Not IntroDone Then frmAbout.UpdateMsg "Starting Opto22 I/O Communications" & vbCrLf
    PowerUpClear_0                     ' PowerUp Clear to All Brains
        Delay
    
    ' Initialize Digitals for base+0 and base+1
    If Not IntroDone Then frmAbout.UpdateMsg "Initialize Opto22 Digital I/O" & vbCrLf
    PowerUpClear_100
        Delay
    
    ' Initialize Analogs for base+2 and base+3
    If Not IntroDone Then frmAbout.UpdateMsg "Initialize Opto22 Analog I/O" & vbCrLf
    PowerUpClear_300
        Delay
      
    ' Finish with setup of all the Watchdog Timers
    If Not IntroDone Then frmAbout.UpdateMsg "Setting Opto22 I/O Watchdog Timers" & vbCrLf
    WatchdogSetup 32
        Delay

End Function

Function Delay()
Dim count As Long
Dim count1 As Long
      For count = 0 To 500000    ' try delaying  some
        count1 = count1 + 1
      Next count
End Function

Function PowerUpClear_100()             ' Digital Modules
'
Dim addr As Integer
Dim baseaddr, baseslot As Integer
Dim chan As Integer
Dim slot As Integer
Dim stn As Integer
Dim indx1, indx2, indx3, indx4 As Integer
Dim maxindx As Integer
Dim tempmax As Integer
Dim pos0 As Long
Dim ErrTxt$
Dim count As Long

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 100

    ' Is a FID Board on this system?
'    If USINGFIDANALYZER And FIDBoard <> 0 Then
'        maxindx = 10
'    Else
        maxindx = 9
'    End If
    
    '  OPTO Module Information based on a 12 slot board
    For indx1 = 0 To maxindx
        If (indx1 = 10 Or indx1 <= NR_STN) And (Node_Info(indx1) > 0) Then
            ' Determine the Base Addresses
            Select Case indx1
                Case 0          ' Main Board
                    baseaddr = 0
                Case 1 To 9     ' Station Boards
                    baseaddr = indx1 * 4
                Case 10         ' FID Board
                    baseaddr = 10
            End Select
            For indx2 = 0 To 1
                addr = baseaddr + indx2
                pos0 = 0
                ' Determine the Base Slot for the current address
                baseslot = indx2 * 4
                ' Look for Digital Modules          1=DI, 2=DO
                tempmax = baseslot + 3
                For slot = baseslot To tempmax
                    For indx3 = 0 To 3
                        chan = ((4 * (slot - baseslot)) + indx3)
                        Select Case Opto_Info(baseaddr, slot)
                            Case optotypeDI      ' DI
                                editSend(chan) = Val(0)     '  0 for DI's
                                pos0 = pos0 + 2 ^ chan
                            Case optotypeDO      ' DO
                                editSend(chan) = Val(128)   '128 for DO's
                                pos0 = pos0 + 2 ^ chan
                            Case Else
                                editSend(chan) = Val(0)
                        End Select
                    Next indx3
                Next slot
                OptoChanMask(addr) = pos0
                
                Brick_Address = Val(addr)
                Brick_Cmd = Val(100)
                Brick_Pos(0) = Val(pos0)
                Brick_Pos(1) = Val(0)
                
                For indx4 = 0 To 15
                    Brick_Send(indx4) = Val(editSend(indx4))
                Next indx4
                
                Brick_Error = SendMIO(Brick_Handle, _
                                      Brick_Address, _
                                      Brick_Cmd, _
                                      Brick_Pos(0), _
                                      Brick_Send(0), _
                                      Brick_Rece(0))
                  
                Delay
                ErrTxt$ = MisticError(Brick_Error)
                For indx4 = 0 To 15
                    DoEvents
                Next indx4
                If Brick_Error < "0" Then           ' Need to add error logging some day
                    ErrTxt$ = MisticError(Brick_Error)
                    ErrorLbl.Caption = str$(Brick_Error) + "   " + ErrTxt$
                    Opto_Error_data = str$(Brick_Error) + "   " + ErrTxt$
                    Delay_Box ErrTxt$, MSGDELAY, msgSHOW        '+ " opto 100", 2
                    Write_ELog ErrTxt$                          '+ " opto 100"
                    Exit Function
                Else
                    ErrorLbl.Caption = "none"
                End If
                
                For indx4 = 0 To 15
                    editRece(indx4).text = str$(Brick_Rece(indx4))
                    Opto_Rec_Data(indx4) = str$(Brick_Rece(indx4))
                Next indx4
                btnClear_Click
        
                
            Next indx2
        End If
    Next indx1

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

Function PowerUpClear_300()             ' Analog Modules
'
Dim addr As Integer
Dim baseaddr As Integer
Dim basechan As Integer
Dim baseslot As Integer
Dim chan As Integer
Dim slot As Integer
Dim stn As Integer
Dim indx1, indx2, indx3, indx4 As Integer
Dim maxindx As Integer
Dim tempmax As Integer
Dim pos0 As Long
Dim ErrTxt$
Dim count As Long

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 300

    ' Is a FID Board on this system?
'    If USINGFIDANALYZER And FIDBoard <> 0 Then
'        maxindx = 10
'    Else
        maxindx = 9
'    End If
    
    '  OPTO Module Information based on a 12 slot board
    '  OPTO Module Information
    For indx1 = 0 To maxindx
        If (indx1 = 10 Or indx1 <= NR_STN) And (Node_Info(indx1) > 0) Then
            ' Determine the Base Addresses
            Select Case indx1
                Case 0          ' Main Board
                    baseaddr = 0
                Case 1 To 9     ' Station Boards
                    baseaddr = indx1 * 4
                Case 10         ' FID Board
                    baseaddr = 10
            End Select
            For indx2 = 2 To 3
                If ((indx2 = 2) Or (Node_Info(indx1) <> 8)) Then
                
                    ' Determine the address of current node
                    addr = baseaddr + indx2
                    pos0 = 0
                    ' Determine the Base Slot for the current node
     '               Select Case indx2
     '                   Case 2
'    '                        baseslot = 6
                            baseslot = (indx2 - 2) * 8
     '                   Case 3
     '                       baseslot = 8
     '               End Select
                    ' Determine the Max Slot and the First Channel for the current node
                    tempmax = baseslot + 7
'                    If indx2 = 2 Then
                        ' Base + 2 = chan 12-15 Only;  2 chan/slot
                        ' Base + 2 = chan 0-15;  2 chan/slot
'                        tempmax = baseslot + 1
'                        basechan = 12
'                        basechan = 0
'                    ElseIf indx2 = 3 Then
                        ' Base + 3 = no channels;                   -  8 slot board
                        ' Base + 3 = chan 0-7 Only;  2 chan/slot    - 12 slot board
                        ' Base + 3 = chan 0-15;  2 chan/slot        - 16 slot board
'                        Select Case Node_Info(indx1)
'                            Case 8
'                                tempmax = 0
'                            Case 12, 16
'                                tempmax = baseslot + 1
'                        End Select
                        basechan = 0
'                    End If
                    ' Look for Analog Modules          3=AI, 4=AO, 5=TC TypeJ, 6=TC Type K
                    ' *******************************************************************************
                    '     0   Not used
                    '     5   Type J thermocouple                                   (hex 05)
                    '     7   0 to 10 volt dc (not preferred one)                   (hex 07)
                    '     8   Type K thermocouple                                   (hex 08)
                    '    10   100 Ohm RTD                                           (hex 0A)
                    '    12   +/- 10 volt Blue module AI inputs                     (hex 0C)
                    '   165   0 to 10 volt dc (Preferred) Green module AO outputs   (hex A5)
                    ' *******************************************************************************
                    For slot = baseslot To tempmax
                        For indx3 = 0 To 1
                            chan = ((2 * (slot - baseslot)) + basechan + indx3)
                            Select Case Opto_Info(baseaddr, slot)
                                Case optotypeAI       ' AI
                                    editSend(chan) = Val("12")
                                    pos0 = pos0 + 2 ^ chan
                                Case optotypeAO       ' AO
                                    editSend(chan) = Val("165")
                                    pos0 = pos0 + 2 ^ chan
                                Case optotypeTcJ      ' TC Type J
                                    editSend(chan) = Val("5")
                                    pos0 = pos0 + 2 ^ chan
                                Case optotypeTcK      ' TC Type K
                                    editSend(chan) = Val("8")
                                    pos0 = pos0 + 2 ^ chan
                                Case optotypeRTD      ' RTD 100 Ohm
                                    editSend(chan) = Val("10")
                                    pos0 = pos0 + 2 ^ chan
                                Case Else
                                    editSend(chan) = Val("0")
                            End Select
                        Next indx3
                    Next slot
                    
                    ' Setup the command's parameters
                    Brick_Address = Val(addr)
                    Brick_Cmd = Val(300)
                    Brick_Pos(0) = Val(pos0)
                    Brick_Pos(1) = Val(0)
                    For indx4 = 0 To 15
                        Brick_Send(indx4) = Val(editSend(indx4))
                    Next indx4
                    
                    ' Send the command
                    Brick_Error = SendMIO(Brick_Handle, _
                                          Brick_Address, _
                                          Brick_Cmd, _
                                          Brick_Pos(0), _
                                          Brick_Send(0), _
                                          Brick_Rece(0))
                      
                    Delay
                    
                    ' Check for errors
                    ErrTxt$ = MisticError(Brick_Error)
                    For indx4 = 0 To 15
                        DoEvents
                    Next indx4
                    If Brick_Error < "0" Then           ' Need to add error logging some day
                        ErrTxt$ = MisticError(Brick_Error)
                        ErrorLbl.Caption = str$(Brick_Error) + "   " + ErrTxt$
                        Opto_Error_data = str$(Brick_Error) + "   " + ErrTxt$
                        Delay_Box ErrTxt$, MSGDELAY, msgSHOW        '+ " opto 100", 2
                        Write_ELog ErrTxt$                          '+ " opto 100"
                        Exit Function
                    Else
                        ErrorLbl.Caption = "none"
                    End If
                    
                    ' Copy the returned data
                    For indx4 = 0 To 15
                        editRece(indx4).text = str$(Brick_Rece(indx4))
                        Opto_Rec_Data(indx4) = str$(Brick_Rece(indx4))
                    Next indx4
                    btnClear_Click
            
                End If
            Next indx2
        End If
    Next indx1

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

Function PowerUpClear_0()               'General clear for all B3000 brains
'
Dim iii As Integer                       'All base addresses each OPTO22 0 through 3
Dim ii As Integer                        'Even though they may not be present
Dim i As Integer
Dim ErrTxt$

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 5

    i = 0
    ii = 0
    iii = 0
    
    For iii = 0 To NR_STN                                   ' board baseaddress
        If Node_Info(iii) > 0 Then
            For ii = 0 To 3                                 ' additional addresses
                editAddress = (4 * iii) + ii
                editCommand = Val(0)
                editPos(1) = Val(0)
                editPos(0) = Val(0)
            
                Brick_Address = Val(editAddress)
                Brick_Cmd = Val(editCommand)                ' Command = 0 for brick Power up/clear
                Brick_Pos(0) = Val(editPos(0))
                Brick_Pos(1) = Val(editPos(1))
                For i = 0 To 15                             ' channels 0 to 15
                    editRece(i).text = "0"                  ' zero out values
                    editSend(i) = Val(0)
                Next i
            
                For i = 0 To 15
                    Brick_Send(i) = Val(editSend(i))
                Next i
               Brick_Error = SendMIO(Brick_Handle, _
                                      Brick_Address, _
                                      Brick_Cmd, _
                                      Brick_Pos(0), _
                                      Brick_Send(0), _
                                      Brick_Rece(0))
                Delay
                ErrTxt$ = MisticError(Brick_Error)
                DoEvents
                For i = 0 To 15
                  editRece(i).text = str$(Brick_Rece(i))
                  Opto_Rec_Data(i) = str$(Brick_Rece(i))
                Next i
            Next ii
        End If
    Next iii
    
    btnClear_Click

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

Function PowerUpClear_3()                   'Save Config to EEPROM for all B3000 brains
'
Dim iii As Integer                          'All base addresses each OPTO22 0 through 3
Dim ii As Integer                           'Even though they may not be present
Dim i As Integer
Dim ErrTxt$

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 93, 53

    i = 0
    ii = 0
    iii = 0
    
    For iii = 0 To NR_STN                   ' board baseaddress
        If Node_Info(iii) > 0 Then
            For ii = 0 To 3                     ' additional addresses
                editAddress = (4 * iii) + ii
                editAddress = ii
                editCommand = Val(3)
                editPos(1) = Val(0)
                editPos(0) = Val(0)
    
                Brick_Address = Val(editAddress)
                Brick_Cmd = Val(editCommand)            ' Command = 3 for brick save to EEPROM
                Brick_Pos(0) = Val(editPos(0))
                Brick_Pos(1) = Val(editPos(1))
                For i = 0 To 15                          ' channels 0 to 15
                    editRece(i).text = "0"                   'zero out values
                    editSend(i) = Val(0)
                Next i
    
                For i = 0 To 15
                    Brick_Send(i) = Val(editSend(i))
                Next i
                Brick_Error = SendMIO(Brick_Handle, _
                                      Brick_Address, _
                                      Brick_Cmd, _
                                      Brick_Pos(0), _
                                      Brick_Send(0), _
                                      Brick_Rece(0))
                Delay
                ErrTxt$ = MisticError(Brick_Error)
                DoEvents
                For i = 0 To 15
                    editRece(i).text = str$(Brick_Rece(i))
                    Opto_Rec_Data(i) = str$(Brick_Rece(i))
                Next i
            Next ii
        End If
    Next iii
    
    btnClear_Click

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





