VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmCourses 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Sequence"
   ClientHeight    =   11115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14835
   Icon            =   "frmCourses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   14835
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlLineVolume 
      Height          =   2160
      Left            =   8640
      TabIndex        =   49
      Top             =   840
      Width           =   4245
      _Version        =   65536
      _ExtentX        =   7488
      _ExtentY        =   3810
      _StockProps     =   15
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Alignment       =   6
      Begin VB.CommandButton cmdAcceptLineVolume 
         DisabledPicture =   "frmCourses.frx":57E2
         DownPicture     =   "frmCourses.frx":5EE4
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   77
         ToolTipText     =   "Accept Line Volume Values"
         Top             =   120
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Frame frmLineVolume 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1440
         Left            =   120
         TabIndex        =   51
         Top             =   615
         Visible         =   0   'False
         Width           =   4005
         Begin VB.TextBox txtIDVent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1785
            MaxLength       =   5
            TabIndex        =   60
            Text            =   "0.0"
            ToolTipText     =   "VENT Inside Diameters in inches"
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtIDPurge 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1785
            MaxLength       =   5
            TabIndex        =   59
            Text            =   "0.0"
            ToolTipText     =   "PURGE Inside Diameters in inches"
            Top             =   770
            Width           =   615
         End
         Begin VB.TextBox txtIDLoad 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   1785
            MaxLength       =   5
            TabIndex        =   58
            Text            =   "0.0"
            ToolTipText     =   "LOAD Inside Diameters in inches"
            Top             =   500
            Width           =   615
         End
         Begin VB.TextBox txtVentL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   57
            Text            =   "0.0"
            ToolTipText     =   "Vent Length in feet"
            Top             =   1050
            Width           =   800
         End
         Begin VB.TextBox txtPurgeL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   56
            Text            =   "0.0"
            ToolTipText     =   "Purge Length in feet"
            Top             =   770
            Width           =   800
         End
         Begin VB.TextBox txtLoadL 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
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
            Left            =   2760
            MaxLength       =   5
            TabIndex        =   55
            Text            =   "0.0"
            ToolTipText     =   "Load Length in feet"
            Top             =   500
            Width           =   800
         End
         Begin VB.TextBox txtVentV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   625
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   54
            Text            =   "0.0"
            ToolTipText     =   "Calculated Volume in liters"
            Top             =   1050
            Width           =   615
         End
         Begin VB.TextBox txtPurgeV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   625
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   53
            Text            =   "0.0"
            ToolTipText     =   "Calculated Volume in liters"
            Top             =   770
            Width           =   615
         End
         Begin VB.TextBox txtLoadV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   285
            Left            =   625
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   52
            Text            =   "0.0"
            ToolTipText     =   "Calculated Volume in liters"
            Top             =   500
            Width           =   615
         End
         Begin VB.Label lblLoadDesc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Load"
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
            Left            =   90
            TabIndex        =   75
            Top             =   525
            Width           =   500
         End
         Begin VB.Label lblPurgeDesc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Purge"
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
            Left            =   90
            TabIndex        =   74
            Top             =   795
            Width           =   500
         End
         Begin VB.Label lblVentDesc 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Vent"
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
            Left            =   90
            TabIndex        =   73
            Top             =   1065
            Width           =   500
         End
         Begin VB.Label lblVentL 
            BackStyle       =   0  'Transparent
            Caption         =   "ft."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3630
            TabIndex        =   72
            Top             =   1085
            Width           =   285
         End
         Begin VB.Label lblPurgeL 
            BackStyle       =   0  'Transparent
            Caption         =   "ft."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3630
            TabIndex        =   71
            Top             =   800
            Width           =   285
         End
         Begin VB.Label lblLoadL 
            BackStyle       =   0  'Transparent
            Caption         =   "ft."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3630
            TabIndex        =   70
            Top             =   530
            Width           =   285
         End
         Begin VB.Label lblIDVent 
            BackStyle       =   0  'Transparent
            Caption         =   "in."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2430
            TabIndex        =   69
            Top             =   1065
            Width           =   285
         End
         Begin VB.Label lblIDPurge 
            BackStyle       =   0  'Transparent
            Caption         =   "in."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2430
            TabIndex        =   68
            Top             =   795
            Width           =   285
         End
         Begin VB.Label lblIDLoad 
            BackStyle       =   0  'Transparent
            Caption         =   "in."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2430
            TabIndex        =   67
            Top             =   525
            Width           =   285
         End
         Begin VB.Label lblLineLength 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Length"
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
            Left            =   2760
            TabIndex        =   66
            Top             =   280
            Width           =   1100
         End
         Begin VB.Label lblLineID 
            BackStyle       =   0  'Transparent
            Caption         =   "Line ID"
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
            Left            =   1785
            TabIndex        =   65
            Top             =   285
            Width           =   795
         End
         Begin VB.Label lblVolume 
            BackStyle       =   0  'Transparent
            Caption         =   "Volume"
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
            Left            =   625
            TabIndex        =   64
            Top             =   285
            Width           =   900
         End
         Begin VB.Label lblVentV 
            BackStyle       =   0  'Transparent
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1270
            TabIndex        =   63
            Top             =   1065
            Width           =   455
         End
         Begin VB.Label lblPurgeV 
            BackStyle       =   0  'Transparent
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1270
            TabIndex        =   62
            Top             =   795
            Width           =   455
         End
         Begin VB.Label lblLoadV 
            BackStyle       =   0  'Transparent
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1270
            TabIndex        =   61
            Top             =   525
            Width           =   455
         End
      End
      Begin VB.CommandButton cmdCancelLineVolume 
         DisabledPicture =   "frmCourses.frx":6CE8
         DownPicture     =   "frmCourses.frx":73EA
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":7AEC
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Done; Close Line Volume"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   480
      End
      Begin VB.Label lblLineVolume 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Line Volume"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   960
         TabIndex        =   76
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame frmSelCourse 
      Appearance      =   0  'Flat
      Caption         =   "Selected Course"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2955
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   14610
      Begin VB.CommandButton cmdSelectRcp 
         Caption         =   "Select Recipe "
         DisabledPicture =   "frmCourses.frx":81EE
         DownPicture     =   "frmCourses.frx":88F0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4028
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":8FF2
         Style           =   1  'Graphical
         TabIndex        =   94
         ToolTipText     =   "Select a Master recipe for this Course"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtRecMsgText 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   9120
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   1470
         Width           =   5250
      End
      Begin VB.TextBox txtRecType 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   1330
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   1470
         Width           =   1050
      End
      Begin VB.CommandButton cmdDnRec 
         Height          =   420
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":96F4
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "next course"
         Top             =   690
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtRecCourse 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   180
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   1470
         Width           =   1050
      End
      Begin VB.CommandButton cmdUpRec 
         Height          =   420
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":9DF6
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "previous course"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtRecPause 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2480
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   1470
         Width           =   1350
      End
      Begin VB.TextBox txtRecRecipe 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   3930
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   1470
         Width           =   1050
      End
      Begin VB.TextBox txtRecCycles 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   5080
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   1470
         Width           =   1050
      End
      Begin VB.TextBox txtRecLoadRate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   6230
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   1470
         Width           =   1350
      End
      Begin VB.TextBox txtRecPurgeRate 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   7680
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1470
         Width           =   1350
      End
      Begin VB.CommandButton cmdNewCourse 
         Caption         =   "New Course "
         DisabledPicture =   "frmCourses.frx":A4F8
         DownPicture     =   "frmCourses.frx":ABFA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":B2FC
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Add a New Course"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update Course "
         DisabledPicture =   "frmCourses.frx":B9FE
         DownPicture     =   "frmCourses.frx":C100
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5940
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":C802
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Update this Course"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete Course "
         DisabledPicture =   "frmCourses.frx":CF04
         DownPicture     =   "frmCourses.frx":D606
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   9840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":DD08
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Delete this Course"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.Label lblCourseMsgText 
         BackStyle       =   0  'Transparent
         Caption         =   "Message Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9120
         TabIndex        =   47
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Label lbPurgeRateDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Rate in"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   6240
         TabIndex        =   44
         Top             =   2520
         Width           =   1605
      End
      Begin VB.Label lbPurgeRateDescUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "splm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   7920
         TabIndex        =   43
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label lblLoadRateDescUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "grams/hr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   7920
         TabIndex        =   42
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblLoadRateDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Load Rate in"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   6240
         TabIndex        =   41
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label lblPauseDurDescUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   7920
         TabIndex        =   40
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label lblPauseDurDesc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pause Duration in"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   6000
         TabIndex        =   39
         Top             =   2040
         Width           =   1845
      End
      Begin VB.Label lblCourse 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Course Type:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblCourseType 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 = Wait for OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Index           =   1
         Left            =   1560
         TabIndex        =   37
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblCourseType 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2 = Pause"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Index           =   2
         Left            =   1560
         TabIndex        =   36
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lblCourseType 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3 = Recipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Index           =   3
         Left            =   1560
         TabIndex        =   35
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label lblCourseNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Course"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   33
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCourseType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1335
         TabIndex        =   32
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCoursePause 
         BackStyle       =   0  'Transparent
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2475
         TabIndex        =   31
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label lblCourseRecipe 
         BackStyle       =   0  'Transparent
         Caption         =   "Recipe"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3930
         TabIndex        =   30
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCourseCycles 
         BackStyle       =   0  'Transparent
         Caption         =   "Cycles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5085
         TabIndex        =   29
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label lblCourseLoadRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6225
         TabIndex        =   28
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblCoursePurgeRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7680
         TabIndex        =   27
         Top             =   1200
         Width           =   1200
      End
   End
   Begin Threed.SSPanel pnlMsg 
      Height          =   1785
      Left            =   60
      TabIndex        =   13
      Top             =   9240
      Width           =   14670
      _Version        =   65536
      _ExtentX        =   25876
      _ExtentY        =   3149
      _StockProps     =   15
      Caption         =   "SSPanel1"
      BackColor       =   14215660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   1575
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   14470
      End
   End
   Begin VB.TextBox txtNotHighlight 
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Text            =   "NOT Highlight"
      Top             =   11400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmNotHighlight 
      Caption         =   "NOT highlight"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   11
      Top             =   11640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmHighlight 
      BackColor       =   &H8000000D&
      Caption         =   "highlight"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   11880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSAdodcLib.Adodc adoCourses 
      Height          =   330
      Left            =   6000
      Top             =   5760
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=CpsRecipes"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "CpsRecipes"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM [zCourses] ORDER BY [CourseNumber] DESC"
      Caption         =   "adoCourses"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dbgCourses 
      Bindings        =   "frmCourses.frx":E40A
      Height          =   7155
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   14610
      _ExtentX        =   25770
      _ExtentY        =   12621
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Courses"
      ColumnCount     =   12
      BeginProperty Column00 
         DataField       =   "Station"
         Caption         =   "Station"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Shift"
         Caption         =   "Shift"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CourseNumber"
         Caption         =   "Course"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Type"
         Caption         =   "Type"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "PauseDuration"
         Caption         =   "PauseDuration"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "RecipeNumber"
         Caption         =   "Recipe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Cycles"
         Caption         =   "Cycles"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "LoadRate"
         Caption         =   "LoadRate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "PurgeRate"
         Caption         =   "PurgeRate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "EstCourseDuration"
         Caption         =   "EstCourseDuration"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "MsgText"
         Caption         =   "MsgText"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "SeqNum"
         Caption         =   "SeqNum"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   975.118
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   810.142
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   780.095
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2580.095
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   3974.74
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1454.74
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSeqDescription 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   420
      Left            =   1770
      MaxLength       =   50
      TabIndex        =   6
      Text            =   "description"
      ToolTipText     =   "Enter up to 50 Character Description"
      Top             =   960
      Width           =   6795
   End
   Begin VB.TextBox txtAuxScale 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   420
      Left            =   13905
      MaxLength       =   5
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Aux Scale # for this Job Sequence"
      Top             =   1440
      Width           =   465
   End
   Begin VB.TextBox txtPriScale 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   420
      Left            =   13905
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Primary Scale # for this Job Sequence"
      Top             =   960
      Width           =   465
   End
   Begin VB.PictureBox pbControlBtns 
      Align           =   1  'Align Top
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14835
      TabIndex        =   0
      Top             =   0
      Width           =   14835
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         DisabledPicture =   "frmCourses.frx":E423
         DownPicture     =   "frmCourses.frx":EB25
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":F227
         Style           =   1  'Graphical
         TabIndex        =   95
         ToolTipText     =   "Open Master JobSequence List"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdNewSeq 
         Caption         =   "New Sequence"
         DisabledPicture =   "frmCourses.frx":F929
         DownPicture     =   "frmCourses.frx":1002B
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6705
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1072D
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Add a New JobSequence"
         Top             =   -240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdLoadAll 
         Caption         =   "Import Master "
         DisabledPicture =   "frmCourses.frx":10E2F
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":11531
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "Import a Master Job Sequence"
         Top             =   -240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdLoadCourses 
         Caption         =   "Only Courses "
         DisabledPicture =   "frmCourses.frx":11C33
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4815
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":12335
         Style           =   1  'Graphical
         TabIndex        =   92
         ToolTipText     =   "Import Only the Courses from a Master Job Sequence"
         Top             =   -240
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdLoadDefault 
         Caption         =   " Default Sequence"
         DisabledPicture =   "frmCourses.frx":12A37
         DownPicture     =   "frmCourses.frx":13139
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2760
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1383B
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Set JobSequence to Default Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdValidateSeq 
         Caption         =   "Validate Sequence"
         DisabledPicture =   "frmCourses.frx":13F3D
         DownPicture     =   "frmCourses.frx":1463F
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":14D41
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Validate the JobSequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdEditSeq 
         Caption         =   "Edit Sequence"
         DisabledPicture =   "frmCourses.frx":15443
         DownPicture     =   "frmCourses.frx":15B45
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   9240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":16247
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Begin Editing the JobSequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSaveSeq 
         Caption         =   "Save Sequence"
         DisabledPicture =   "frmCourses.frx":16949
         DownPicture     =   "frmCourses.frx":1704B
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   11160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1774D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Save the JobSequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDn 
         Caption         =   "Prev"
         DisabledPicture =   "frmCourses.frx":17E4F
         DownPicture     =   "frmCourses.frx":18551
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4815
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":18C53
         Style           =   1  'Graphical
         TabIndex        =   87
         ToolTipText     =   "Previous Master Job Sequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdPgDn 
         Caption         =   "Pg Prev"
         DisabledPicture =   "frmCourses.frx":19355
         DownPicture     =   "frmCourses.frx":19A57
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1A159
         Style           =   1  'Graphical
         TabIndex        =   86
         ToolTipText     =   "x10 Previous Master Job Sequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Next"
         DisabledPicture =   "frmCourses.frx":1A85B
         DownPicture     =   "frmCourses.frx":1AF5D
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6705
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1B65F
         Style           =   1  'Graphical
         TabIndex        =   85
         ToolTipText     =   "Next Master Job Sequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdPgUp 
         Caption         =   "Pg Next"
         DisabledPicture =   "frmCourses.frx":1BD61
         DownPicture     =   "frmCourses.frx":1C463
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1CB65
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "x10 Next Master Job Sequence"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdRestoreSeq 
         Caption         =   " Restore Sequence"
         DisabledPicture =   "frmCourses.frx":1D267
         DownPicture     =   "frmCourses.frx":1D969
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   12720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1E06B
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Reload JobSequence Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         DisabledPicture =   "frmCourses.frx":1E76D
         DownPicture     =   "frmCourses.frx":1EE6F
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":1F571
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Paste Job Sequence Values from the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         DisabledPicture =   "frmCourses.frx":1FC73
         DownPicture     =   "frmCourses.frx":20375
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":20A77
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Copy Job Sequence Values to the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdLineVolume 
         Caption         =   "Line Volume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   13920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCourses.frx":21179
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Show Job Sequence Line Volume Settings"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin Threed.SSPanel txtDispSeqNum 
         Height          =   840
         Left            =   5670
         TabIndex        =   88
         ToolTipText     =   "Click for list of Defined Job Sequences"
         Top             =   0
         Width           =   1035
         _Version        =   65536
         _ExtentX        =   1826
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "01"
         ForeColor       =   -2147483646
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   24.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
      Begin VB.Label lblStnDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "station shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   96
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
      End
   End
   Begin VB.Label lblCourses 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "888"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      DataMember      =   ".Recordset.RecordCount"
      DataSource      =   "adoCourses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1770
      TabIndex        =   79
      ToolTipText     =   "Number of Courses in this Job Sequence"
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label lblSeqDuration 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8 hours  88 minutes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4590
      TabIndex        =   78
      ToolTipText     =   "Duration of this Job Sequence"
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label lblSeqDurationDesc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3120
      TabIndex        =   48
      Top             =   1485
      Width           =   1335
   End
   Begin VB.Label lblNumCourses 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Courses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   8
      Top             =   1485
      Width           =   1335
   End
   Begin VB.Label lblDescription 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   1005
      Width           =   1335
   End
   Begin VB.Label lblAuxScale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aux. Scale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12240
      TabIndex        =   5
      Top             =   1485
      Width           =   1605
   End
   Begin VB.Label lblPriScale 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Primary Scale"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12240
      TabIndex        =   3
      Top             =   1005
      Width           =   1605
   End
End
Attribute VB_Name = "frmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' module 390 ******* Job Sequence Screen
Option Explicit
'
' NOTES
'   1. All Editing is done in New
'   2. All Saves are from New
'   3. Saves include Copy from New to Curr & Prev
'
Private JobSeqMode As Integer                       ' 0=master; 1=station
Private DispSeqNum As Integer                       ' Current Master Sequence index
Private ScreenBkgdColor As Long
Private ScreenDescription As String
Private ScreenDispFlag As Boolean
Private StnShftDescription As String
Private dbDbase As Database
Private rsRecord  As Recordset
Private Criteria As String
Private SelectedCourse As Integer                   ' currently selected course number
Private CourseCount As Integer                      ' number of courses in the JobSequence
Private tempSeq As JobSequence                      ' temp. JobSequence; used for Resequencing
Private MemSeq As JobSequence                       ' clipboard for Cut/Paste
Private NewSeq As JobSequence                       ' JobSequence being edited
Private CurrSeq As JobSequence                      ' "current" JobSequence
Private PrevSeq As JobSequence                      ' copy of last "Saved" JobSequence
Private bUnUpdated As Boolean                       ' Flag - Whether unUpdated (New) JobSequence Course changes exist
Private bUnValidSeq As Boolean                      ' Flag - Whether unValidated (New) JobSequence changes exist
Private bUnSavedSeq As Boolean                      ' Flag - Whether unSaved (Curr) JobSequence changes exist
Private bEmptyMemSeq As Boolean                     ' Flag - Whether the MemSeq (i.e. clipboard) is Empty
Private bEditing As Boolean                         ' Flag - Editing(New) Or Viewing(Curr) the JobSequence
Private bLoadCoursesOnly As Boolean                 ' Flag - Whether to Only Load Courses when Loading a MasterSequence into a StationSequence
Private Dbg_EditHeight As Integer
Private Dbg_FullHeight As Integer

'Private Sub adoCourses_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    dbgCourses.Refresh
'End Sub

Public Sub ChgJobSeqMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    JobSeqMode = IIf((NewMode = MASTERMODE Or NewMode = STATIONMODE), NewMode, MASTERMODE)
    Select Case JobSeqMode
        Case MASTERMODE
            ' screen description
            ScreenDescription = "Master Job Sequence Properties"
            ' screen background color
            ScreenBkgdColor = MasterMode_BackColor
            ' show items
            ScreenDispFlag = True
        Case STATIONMODE
            ' station/shift description
            StnShftDescription = "Station #" & Format(DispStn, "#0")
            If NR_SHIFT > 1 Then StnShftDescription = StnShftDescription & "  Shift #" & Format(DispShift, "#0")
            StnShftDescription = StnShftDescription & "  Job Sequence Properties"
            ' screen description
            ScreenDescription = StnShftDescription
            ' screen background color
            ScreenBkgdColor = StationMode_BackColor
            ' hide items
            ScreenDispFlag = False
    End Select
    ' screen description
    frmCourses.Caption = ScreenDescription
    ' set screen background colors
    frmCourses.BackColor = ScreenBkgdColor
    pbControlBtns.BackColor = ScreenBkgdColor
    txtDispSeqNum.BackColor = ScreenBkgdColor
    ' show items ??
    SetButtonsForMode JobSeqMode
End Sub

Public Sub LoadRcpNum(ByVal NewRcp As Integer)
    txtRecRecipe.text = Format(NewRcp, "##0")
End Sub

Private Sub ExitScreen()
    ' close Sequence/Recipe/Canister database
    dbDbase.Close
    ' unload form
    Unload Me
    Set frmCourses = Nothing
End Sub

Private Sub SetSeqToDefault(iSeq As JobSequence)
'
'
Dim iCourse As Integer

    ' Clear New Courses
    ClearCourses iSeq
    Select Case JobSeqMode
        Case MASTERMODE
            ' Set Master Sequence Information
            iSeq.Number = DispSeqNum
            iSeq.Description = "default master sequence #" & Format(DispSeqNum, "##0")
            iSeq.PriScaleNo = CInt(0)
            iSeq.AuxScaleNo = CInt(0)
            iSeq.EstSeqDuration = 0
            iSeq.EstSeqDurDesc = "undefined"
        Case STATIONMODE
            ' Set Station Sequence Information
            iSeq.Number = CInt(0)
            iSeq.Description = "default station sequence"
            iSeq.PriScaleNo = STN_INFO(DispStn).DefPriScale
            iSeq.AuxScaleNo = STN_INFO(DispStn).DefAuxScale
            iSeq.EstSeqDuration = EstimatedRcpDuration(StationRecipe(DispStn, DispShift), StationCanister(DispStn, DispShift), StationProfile(DispStn, DispShift))
            iSeq.EstSeqDurDesc = DurationDescription(iSeq.EstSeqDuration)
    End Select
    iSeq.NumCourses = 1
    iSeq.IDLoad = 0
    iSeq.IDPurge = 0
    iSeq.IDVent = 0
    iSeq.LoadL = 0
    iSeq.LoadV = 0
    iSeq.PurgeL = 0
    iSeq.PurgeV = 0
    iSeq.VentL = 0
    iSeq.VentV = 0
    iSeq.Validated = True
    For iCourse = 1 To MAX_COURSES
        iSeq.CourseData(iCourse).CourseNumber = iCourse
        iSeq.CourseData(iCourse).Type = IIf((iCourse = 1), courseRecipe, courseUndefined)
        iSeq.CourseData(iCourse).EstCourseDuration = IIf((iCourse = 1), iSeq.EstSeqDuration, 0)
        iSeq.CourseData(iCourse).RecipeNumber = 0
        iSeq.CourseData(iCourse).Cycles = 0
        iSeq.CourseData(iCourse).LoadRate = 0
        iSeq.CourseData(iCourse).MsgText = "none"
        iSeq.CourseData(iCourse).PauseDuration = 0
        iSeq.CourseData(iCourse).PurgeRate = 0
    Next iCourse
    bUnValidSeq = True
    SelectedCourse = 1
End Sub

Private Sub SaveMasterSequence(ByVal iMaster As Integer)
Dim iCourse As Integer

        ' Locate Master Sequence Information Record
        Criteria = "SELECT * FROM [MasterSequence] WHERE [Number] = " & iMaster & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        If rsRecord.BOF Then
            rsRecord.AddNew
            rsRecord("Number") = iMaster
        Else
          rsRecord.MoveFirst
          rsRecord.Edit
        End If
           
        ' Update Master Sequence Information Record
        rsRecord("Description") = NewSeq.Description
        rsRecord("Courses") = NewSeq.NumCourses
        rsRecord("PriScale") = NewSeq.PriScaleNo
        rsRecord("AuxScale") = NewSeq.AuxScaleNo
        rsRecord("IDLoad") = NewSeq.IDLoad
        rsRecord("IDPurge") = NewSeq.IDPurge
        rsRecord("IDVent") = NewSeq.IDVent
        rsRecord("LoadL") = NewSeq.LoadL
        rsRecord("LoadV") = NewSeq.LoadV
        rsRecord("PurgeL") = NewSeq.PurgeL
        rsRecord("PurgeV") = NewSeq.PurgeV
        rsRecord("VentL") = NewSeq.VentL
        rsRecord("VentV") = NewSeq.VentV
        rsRecord("Validated") = NewSeq.Validated
        rsRecord("EstSeqDuration") = NewSeq.EstSeqDuration
        rsRecord("EstSeqDurDesc") = NewSeq.EstSeqDurDesc
        rsRecord.Update
        rsRecord.Close

        ' Clear Master Sequence Course Information Records
        Criteria = "SELECT * FROM [MasterSequenceCourses] WHERE [SeqNum] = " & iMaster & " "
        Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        If Not rsRecord.BOF Then
            rsRecord.MoveFirst
            rsRecord.MoveLast
            Do While Not rsRecord.BOF
                rsRecord.Delete
                rsRecord.MovePrevious
            Loop
        End If
        For iCourse = 1 To MAX_COURSES
            If NewSeq.CourseData(iCourse).Type <> courseUndefined Then
                rsRecord.AddNew
                rsRecord("SeqNum") = iMaster
                rsRecord("CourseNumber") = NewSeq.CourseData(iCourse).CourseNumber
                rsRecord("Type") = NewSeq.CourseData(iCourse).Type
                rsRecord("PauseDuration") = NewSeq.CourseData(iCourse).PauseDuration
                rsRecord("RecipeNumber") = NewSeq.CourseData(iCourse).RecipeNumber
                rsRecord("Cycles") = NewSeq.CourseData(iCourse).Cycles
                rsRecord("LoadRate") = NewSeq.CourseData(iCourse).LoadRate
                rsRecord("PurgeRate") = NewSeq.CourseData(iCourse).PurgeRate
                rsRecord("EstCourseDuration") = NewSeq.CourseData(iCourse).EstCourseDuration
                rsRecord("MsgText") = NewSeq.CourseData(iCourse).MsgText
                rsRecord.Update
            End If
        Next iCourse
        rsRecord.Close
End Sub

Public Sub InitSeqRcp()
    JobSeqAutoEdit = True
    Select Case JobSeqMode
        Case MASTERMODE
            ' master
            If ((DispSeqNum < 1) Or (DispSeqNum > NR_JOBSEQ)) Then DispSeqNum = 1
            CopyMasterToCurr DispSeqNum
        Case STATIONMODE
            ' station
            If ((StationSequence(DispStn, DispShift).Number < 0) _
             Or (StationSequence(DispStn, DispShift).Number > NR_JOBSEQ)) Then
               DispSeqNum = 0
            Else
               DispSeqNum = StationSequence(DispStn, DispShift).Number
            End If
            CopyStationToCurr DispStn, DispShift
    End Select
    PrevSeq = CurrSeq
    DisplaySequence CurrSeq
    JobSeqAutoEdit = False
    ' update flags
    bUnSavedSeq = False
    bUnValidSeq = False
    bUnUpdated = False
    SetupForEdit False
End Sub

Private Function ValidSequence() As Boolean
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 50, 111
Dim flag As Boolean
Dim Message As String

    flag = True
    Message = ""
    
    If Len(txtSeqDescription.text) < 1 Then txtSeqDescription.text = " "
    Select Case JobSeqMode
        Case MASTERMODE
            ' Scale assignments
            If Not Range_Check(txtPriScale, 0, CSng(NR_SCALES), "Primary Scale") Then flag = False
            If Not Range_Check(txtAuxScale, 0, CSng(NR_SCALES), "Aux Scale") Then flag = False
            ' Line Volume values
            If Not IsNumeric(txtIDLoad.text) Then txtIDLoad.text = "0"
            If Not IsNumeric(txtLoadL.text) Then txtLoadL.text = "0"
            If Not IsNumeric(txtIDPurge.text) Then txtIDPurge.text = "0"
            If Not IsNumeric(txtPurgeL.text) Then txtPurgeL.text = "0"
            If Not IsNumeric(txtIDVent.text) Then txtIDVent.text = "0"
            If Not IsNumeric(txtVentL.text) Then txtVentL.text = "0"
            txtLoadV.text = "0"
            txtPurgeV.text = "0"
            txtVentV.text = "0"
        Case STATIONMODE
            ' Scale assignments
            If USINGHARDPIPEDSCALES Then
                ' two scales per station, fixed assignments for pri & aux for each station; stn#1 pri = 1, stn#1 aux = 2, etc.
                If (CInt(txtPriScale.text) <> STN_INFO(DispStn).DefPriScale) Then
                    txtPriScale.text = Format(STN_INFO(DispStn).DefPriScale, "#0")
                    Message = Message & "Corrected Primary Scale Assignment" & vbCrLf
                End If
                If (CInt(txtAuxScale.text) <> STN_INFO(DispStn).DefAuxScale) Then
                    txtAuxScale.text = Format(STN_INFO(DispStn).DefAuxScale, "#0")
                    Message = Message & "Corrected Aux Scale Assignment" & vbCrLf
                End If
                If (Len(Message) > 1) Then lblMessage.Caption = lblMessage.Caption & Message
            Else
                If Not Range_Check(txtPriScale, 0, CSng(NR_SCALES), "Primary Scale") Then flag = False
                If Not Range_Check(txtAuxScale, 0, CSng(NR_SCALES), "Aux Scale") Then flag = False
            End If
            ' Line Volume values
            If USINGLINEVOLUME Then
                If Not Range_Check(txtIDLoad, 0, 1, "Load Line ID") Then flag = False
                If Not Range_Check(txtIDPurge, 0, 1, "Purge Line ID") Then flag = False
                If Not Range_Check(txtIDVent, 0, 1, "Vent Line ID") Then flag = False
                If Not Range_Check(txtLoadL, 0, 200, "Load Line Length") Then flag = False
                If Not Range_Check(txtPurgeL, 0, 200, "Purge Line Length") Then flag = False
                If Not Range_Check(txtVentL, 0, 200, "Vent Line Length") Then flag = False
                If flag Then
                     txtLoadV = Format(LineVolume(CSng(txtIDLoad), CSng(txtLoadL)), "00.00")
                     txtPurgeV = Format(LineVolume(CSng(txtIDPurge), CSng(txtPurgeL)), "00.00")
                     txtVentV = Format(LineVolume(CSng(txtIDVent), CSng(txtVentL)), "00.00")
                End If
            Else
                If Not IsNumeric(txtIDLoad.text) Then txtIDLoad.text = "0"
                If Not IsNumeric(txtLoadL.text) Then txtLoadL.text = "0"
                If Not IsNumeric(txtIDPurge.text) Then txtIDPurge.text = "0"
                If Not IsNumeric(txtPurgeL.text) Then txtPurgeL.text = "0"
                If Not IsNumeric(txtIDVent.text) Then txtIDVent.text = "0"
                If Not IsNumeric(txtVentL.text) Then txtVentL.text = "0"
                txtLoadV.text = "0"
                txtPurgeV.text = "0"
                txtVentV.text = "0"
            End If
    End Select
        
    ValidSequence = flag
    
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

Private Sub CourseDataHasChanged()
    If Not JobSeqAutoEdit Then
        bUnUpdated = True
        If JobSeqMode = STATIONMODE Then DispSeqNum = 0
        SetupForEdit bEditing
    End If
End Sub

Private Sub SeqHasChanged()
    If Not JobSeqAutoEdit Then
        bUnValidSeq = True
        bUnSavedSeq = False
        If JobSeqMode = STATIONMODE Then DispSeqNum = 0
        SetupForEdit bEditing
    End If
End Sub

Private Sub cmdAcceptLineVolume_Click()
    pnlLineVolume.Top = OutOfSight
End Sub

Private Sub cmdCancelLineVolume_Click()
    pnlLineVolume.Top = OutOfSight
End Sub

Private Sub cmdCopy_Click()
    lblMessage.Caption = vbCrLf
    Select Case bEditing
        Case True
            MemSeq = NewSeq
            bEmptyMemSeq = False
        Case False
            ' nothing to do
    End Select
    SetupForEdit bEditing
End Sub

Private Function ValidCourses(iSeq As JobSequence) As Boolean
Dim cumDuration As Single
Dim iCourse As Integer
Dim flag As Integer
Dim multCalcWcFlag As Boolean
Dim calcWcCourse As Integer
Dim firstRcpCourse As Integer
    
    flag = 0
    cumDuration = 0
    multCalcWcFlag = False
    calcWcCourse = 0
    firstRcpCourse = 0
    ' Check each recipe for the station
    If JobSeqMode = STATIONMODE Then
        If StationCanister(DispStn, DispShift).Validated Then
            frmRecipe.Show
            frmRecipe.tmrUpdate.Enabled = True
            frmRecipe.ChgRecipeMode STATIONMODE
            For iCourse = 1 To iSeq.NumCourses
                If flag = 0 Then
                    Select Case iSeq.CourseData(iCourse).Type
                        Case courseWait
                            ' Wait for operator OK
                            iSeq.CourseData(iCourse).OkToProceed = False
                            iSeq.CourseData(iCourse).EstCourseDuration = 15
                        Case coursePause
                            ' Pause for x minutes
                            If iSeq.CourseData(iCourse).PauseDuration < 0 Then flag = iCourse
                            If iSeq.CourseData(iCourse).PauseDuration > 999 Then flag = iCourse
                            If flag = 0 Then iSeq.CourseData(iCourse).EstCourseDuration = iSeq.CourseData(iCourse).PauseDuration
                        Case courseRecipe
                            ' run Recipe x
                            LoadCourseRecipe iSeq, iCourse
                            If Not frmRecipe.OkToRunRecipeInStation Then flag = iCourse
                            If flag = 0 Then
                                If (firstRcpCourse = 0) Then firstRcpCourse = iCourse
                                If (frmRecipe.optUpdateCanWc.Value = cYES) Then
                                    If calcWcCourse <> 0 Then multCalcWcFlag = True
                                    calcWcCourse = iCourse
                                End If
                                frmRecipe.ExportRecipe
                                frmPurgeProfile.Show
                                frmPurgeProfile.ChgProfileMode (MASTERMODE)
                                frmPurgeProfile.LoadNewProf ExportedRecipe.Purge_ProfileNumber
                                frmPurgeProfile.ExportProfile
                                iSeq.CourseData(iCourse).EstCourseDuration = EstimatedRcpDuration(ExportedRecipe, StationCanister(DispStn, DispShift), ExportedProfile)
                                Unload frmPurgeProfile
                            End If
                        Case Else
                            ' invalid course type
                            flag = iCourse
                    End Select
                End If
'                With adoCourses.Recordset
'                    .Fields("EstCourseDuration") = iSeq.CourseData(iCourse).EstCourseDuration
'                    .Update
'                End With
                cumDuration = cumDuration + iSeq.CourseData(iCourse).EstCourseDuration
            Next iCourse
            iSeq.EstSeqDuration = cumDuration
            iSeq.EstSeqDurDesc = DurationDescription(iSeq.EstSeqDuration)
            ' Check for Errors
            If ((StationCanister(DispStn, DispShift).WorkingCapacity = CSng(0)) And (calcWcCourse = 0)) Then
                ' No CalcWC course and Can BWC = 0
                flag = firstRcpCourse
                lblMessage.Caption = lblMessage.Caption & "Validation Failed on Course #" + Format(flag, "#0") & " for Station #" + Format(DispStn, "0") & " Shift #" + Format(DispShift, "0")
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Canister has no BWC and No Course uses the CalcWC End Method"
                lblMessage.Caption = lblMessage.Caption & vbCrLf
            ElseIf (multCalcWcFlag) Then
                ' multiple CalcWC courses
                flag = calcWcCourse
                lblMessage.Caption = lblMessage.Caption & "Validation Failed on Course #" + Format(flag, "#0") & " for Station #" + Format(DispStn, "0") & " Shift #" + Format(DispShift, "0")
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "There is more than one Course using the CalcWC End Method"
                lblMessage.Caption = lblMessage.Caption & vbCrLf
            ElseIf (firstRcpCourse < calcWcCourse) Then
                ' the CalcWC course is Not the first "Recipe" course
                flag = calcWcCourse
                lblMessage.Caption = lblMessage.Caption & "Validation Failed on Course #" + Format(flag, "#0") & " for Station #" + Format(DispStn, "0") & " Shift #" + Format(DispShift, "0")
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "The CalcWC End Method is only allowed for the first Recipe Course"
                lblMessage.Caption = lblMessage.Caption & vbCrLf
            ElseIf flag <> 0 Then
                ' course validation error
                lblMessage.Caption = lblMessage.Caption & "Validation Failed on Course #" + Format(flag, "#0") & " for Station #" + Format(DispStn, "0") & " Shift #" + Format(DispShift, "0")
                lblMessage.Caption = lblMessage.Caption & vbCrLf
            Else
                ' no errors
                lblMessage.Caption = lblMessage.Caption & "All Courses Validated for Station #" + Format(DispStn, "0") & " Shift #" + Format(DispShift, "0")
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                frmRecipe.ExitScreen
            End If
            frmCourses.Show
        Else
            flag = 1
            lblMessage.Caption = lblMessage.Caption & "Must have valid Canister defined first" & vbCrLf
        End If
        StationSequence(DispStn, DispShift).Validated = IIf(flag = 0, True, False)
    End If
    ValidCourses = IIf(flag = 0, True, False)
End Function

Private Sub LoadCourseRecipe(iSeq As JobSequence, ByVal iCourse As Integer)
'
    ' run Recipe x
    Select Case iSeq.CourseData(iCourse).RecipeNumber
        Case 0
            ' run current station recipe with optional changes
            frmRecipe.InitRecipe
            With frmRecipe
                ' optional changes to recipe
                If (iSeq.CourseData(iCourse).Cycles > 0) Then .txtPFCycle.text = Format(iSeq.CourseData(iCourse).Cycles, "##0")
                If (iSeq.CourseData(iCourse).LoadRate > 0) Then .txtLoadRate.text = Format(iSeq.CourseData(iCourse).LoadRate, "##0.0##")
                If (iSeq.CourseData(iCourse).PurgeRate > 0) Then .txtPurgeFlow.text = Format(iSeq.CourseData(iCourse).PurgeRate, "##0.0##")
            End With
        Case Else
            ' run master recipe with changes, some optional
            frmRecipe.LoadNewRcp iSeq.CourseData(iCourse).RecipeNumber
            With frmRecipe
                ' optional changes to recipe
                If (iSeq.CourseData(iCourse).Cycles > 0) Then .txtPFCycle.text = Format(iSeq.CourseData(iCourse).Cycles, "##0")
                If (iSeq.CourseData(iCourse).LoadRate > 0) Then .txtLoadRate.text = Format(iSeq.CourseData(iCourse).LoadRate, "##0.0##")
                If (iSeq.CourseData(iCourse).PurgeRate > 0) Then .txtPurgeFlow.text = Format(iSeq.CourseData(iCourse).PurgeRate, "##0.0##")
                ' job sequence changes to recipe
                .chkPrimaryScale.Value = IIf((iSeq.PriScaleNo > 0), cYES, cNO)
                .chkUseAuxScale = IIf((iSeq.AuxScaleNo > 0), cYES, cNO)
                .txtPrimaryScaleNo.text = Format(iSeq.PriScaleNo, "#0")
                .txtAuxScaleNo.text = Format(iSeq.AuxScaleNo, "#0")
                .txtIDLoad.text = Format(iSeq.IDLoad, "#0.00")
                .txtIDPurge.text = Format(iSeq.IDPurge, "#0.00")
                .txtIDVent.text = Format(iSeq.IDVent, "#0.00")
                .txtLoadL.text = Format(iSeq.LoadL, "##0.00")
                .txtLoadV.text = Format(iSeq.LoadV, "##0.00")
                .txtPurgeL.text = Format(iSeq.PurgeL, "##0.00")
                .txtPurgeV.text = Format(iSeq.PurgeV, "##0.00")
                .txtVentL.text = Format(iSeq.VentL, "##0.00")
                .txtVentV.text = Format(iSeq.VentV, "##0.00")
            End With
    End Select

End Sub

Private Sub cmdEditSeq_Click()
    lblMessage.Caption = vbCrLf
    ' station idle ??
    If ((JobSeqMode = STATIONMODE) And (StationControl(DispStn, DispShift).Mode <> VBIDLE)) Then
        lblMessage.Caption = lblMessage.Caption & "Station Must Be Idle" & vbCrLf
    Else
        Select Case bEditing
            Case False
                ' start editing
                JobSeqAutoEdit = True
                NewSeq = CurrSeq
                ReSeqCourses NewSeq
                DisplaySequence NewSeq
                SelectedCourse = 1
                DisplaySelectedCourse NewSeq
                JobSeqAutoEdit = False
                SetupForEdit True
            Case True
                ' cancel editing
                JobSeqAutoEdit = True
                CurrSeq = PrevSeq
                DisplaySequence CurrSeq
                JobSeqAutoEdit = False
                SetupForEdit False
        End Select
    End If
End Sub

Private Sub cmdLoadAll_Click()
    bLoadCoursesOnly = False
    frmSearchJobSeq.Show
End Sub

Private Sub cmdLoadCourses_Click()
    bLoadCoursesOnly = True
    frmSearchJobSeq.Show
End Sub

Public Sub LoadDefault()
    lblMessage.Caption = vbCrLf
    SetSeqToDefault NewSeq
    DisplaySequence NewSeq
    DisplaySelectedCourse NewSeq
    lblMessage.Caption = lblMessage.Caption & "Default JobSequence Loaded" & vbCrLf
    SetupForEdit True
End Sub

Private Sub cmdLoadDefault_Click()
    LoadDefault
End Sub

Private Sub cmdNewSeq_Click()
Dim iCourse As Integer
    lblMessage.Caption = vbCrLf
    Select Case JobSeqMode
        Case MASTERMODE
            ' Set Master Sequence Information
            NewSeq.Number = DispSeqNum
            NewSeq.Description = "default master sequence #" & Format(DispSeqNum, "##0")
            NewSeq.PriScaleNo = CInt(0)
            NewSeq.AuxScaleNo = CInt(0)
            NewSeq.EstSeqDuration = 0
            NewSeq.EstSeqDurDesc = "undefined"
        Case STATIONMODE
            ' Set Station Sequence Information
            NewSeq.Number = CInt(0)
            NewSeq.Description = "default station sequence"
            NewSeq.PriScaleNo = STN_INFO(DispStn).DefPriScale
            NewSeq.AuxScaleNo = STN_INFO(DispStn).DefAuxScale
            NewSeq.EstSeqDuration = EstimatedRcpDuration(StationRecipe(DispStn, DispShift), StationCanister(DispStn, DispShift), StationProfile(DispStn, DispShift))
            NewSeq.EstSeqDurDesc = DurationDescription(NewSeq.EstSeqDuration)
    End Select
    NewSeq.NumCourses = 1
    NewSeq.IDLoad = 0
    NewSeq.IDPurge = 0
    NewSeq.IDVent = 0
    NewSeq.LoadL = 0
    NewSeq.LoadV = 0
    NewSeq.PurgeL = 0
    NewSeq.PurgeV = 0
    NewSeq.VentL = 0
    NewSeq.VentV = 0
    NewSeq.Validated = False
    For iCourse = 1 To MAX_COURSES
        NewSeq.CourseData(iCourse).CourseNumber = iCourse
        NewSeq.CourseData(iCourse).Type = IIf((iCourse = 1), courseRecipe, courseUndefined)
        NewSeq.CourseData(iCourse).EstCourseDuration = IIf((iCourse = 1), NewSeq.EstSeqDuration, 0)
        NewSeq.CourseData(iCourse).RecipeNumber = 0
        NewSeq.CourseData(iCourse).Cycles = 0
        NewSeq.CourseData(iCourse).LoadRate = 0
        NewSeq.CourseData(iCourse).MsgText = "none"
        NewSeq.CourseData(iCourse).PauseDuration = 0
        NewSeq.CourseData(iCourse).PurgeRate = 0
    Next iCourse
    SelectedCourse = 1
    DisplaySequence NewSeq
    bUnValidSeq = True
    SetupForEdit bEditing
End Sub

Private Sub cmdOpen_Click()
    bLoadCoursesOnly = False
    frmSearchJobSeq.Show
End Sub

Private Sub cmdSelectRcp_Click()
    ' open View Master Recipes screen
    frmSearchRcp.Show
    frmSearchRcp.ChgSelectionDestination rcpdestCourse
End Sub

Public Sub ValidateSeq()
    ' clear user message
    lblMessage.Caption = vbCrLf
    ' reset all backgrounds
    Reset_BackColors
    ' auto-editing begins
    JobSeqAutoEdit = True
    ' screen JobSequence Values OK ??
    If ValidSequence Then
        ' copy screen values to NewSeq
        CopyScreenToSeq NewSeq
        ' copy screen course values to NewSeq
        CopyScreenToCourse NewSeq
        ' check for out-of-sequence courses
        ReSeqCourses NewSeq
        ' courses valid ??
        If ValidCourses(NewSeq) Then
            ' JobSequence is Validated
            NewSeq.Validated = True
            ' update display of JobSequence
            DisplaySequence NewSeq
            ' auto-editing ends
            JobSeqAutoEdit = False
            ' update control flags
            bUnSavedSeq = True
            bUnValidSeq = False
            ' update buttons
            SetupForEdit bEditing
            ' notify user
            lblMessage.Caption = lblMessage.Caption & "Sequence is Valid" & vbCrLf
        Else
            ' auto-editing ends
            JobSeqAutoEdit = False
            ' notify user
            lblMessage.Caption = lblMessage.Caption & "Courses Failed Validation" & vbCrLf
        End If
    Else
        ' auto-editing ends
        JobSeqAutoEdit = False
        ' notify user
        lblMessage.Caption = lblMessage.Caption & "Sequence Failed Validation" & vbCrLf
    End If
End Sub

Private Sub cmdValidateSeq_Click()
    ValidateSeq
End Sub

Private Sub cmdLineVolume_Click()
    pnlLineVolume.Top = 360
End Sub

Private Sub cmdPaste_Click()
    lblMessage.Caption = vbCrLf
    NewSeq = MemSeq
    NewSeq.Number = DispSeqNum
    DisplaySequence NewSeq
    bUnValidSeq = True
    SelectedCourse = 1
    SeqHasChanged
    SetupForEdit bEditing
End Sub

Private Sub cmdDn_Click()
    lblMessage = vbCrLf
    DispSeqNum = IIf(DispSeqNum < 2, NR_JOBSEQ, DispSeqNum - 1)
    LoadMaster DispSeqNum
    PrevSeq = CurrSeq
End Sub

Private Sub cmdUp_Click()
    lblMessage = vbCrLf
    DispSeqNum = IIf(DispSeqNum > NR_JOBSEQ - 1, 1, DispSeqNum + 1)
    LoadMaster DispSeqNum
    PrevSeq = CurrSeq
End Sub

Private Sub cmdPgDn_Click()
    lblMessage = vbCrLf
    DispSeqNum = IIf(DispSeqNum < 12, NR_JOBSEQ, DispSeqNum - 10)
    LoadMaster DispSeqNum
    PrevSeq = CurrSeq
End Sub

Private Sub cmdPgUp_Click()
    lblMessage = vbCrLf
    DispSeqNum = IIf(DispSeqNum > NR_JOBSEQ - 10, 1, DispSeqNum + 10)
    LoadMaster DispSeqNum
    PrevSeq = CurrSeq
End Sub

Public Sub LoadMaster(ByVal iMaster As Integer)
    If ((iMaster < 1) Or (iMaster > NR_JOBSEQ)) Then iMaster = 1
    Select Case JobSeqMode
        Case MASTERMODE
            ' Load a Master to Curr for Viewing
            JobSeqAutoEdit = True
            CopyMasterToCurr iMaster
            DisplaySequence CurrSeq
            JobSeqAutoEdit = False
            ' update flags
            bUnSavedSeq = False
            bUnValidSeq = False
            bUnUpdated = False
            SetupForEdit False
        Case STATIONMODE
            ' Load a Master to New for Editing
            JobSeqAutoEdit = True
            CopyMasterToNew iMaster
            DisplaySequence NewSeq
            SelectedCourse = 1
            DisplaySelectedCourse NewSeq
            JobSeqAutoEdit = False
            ' update flags
            bUnSavedSeq = False
            bUnValidSeq = True
            bUnUpdated = False
            SetupForEdit True
    End Select
End Sub

Private Sub Print_All()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 50, 2
Dim idx As Integer
Dim sdate As String
Dim printstring As String
Dim oldFont As New StdFont
    ' Save current printer font
    oldFont = Printer.Font
    Printer.Font = FILEFONT
    Printer.Font.Size = FILEFONTSIZE
    Printer.Font.Bold = False
    Printer.Font.Italic = False
    
    sdate = "mmm dd, yy  hh:mm"
    Print_Center SysConfig.Heading
    Print_Center SysConfig.Heading2
    Print_Center "CANISTER PRECONDITIONING SYSTEM"
    Print_Line ""
    Print_Center ("Job Sequence Listing")
    Print_Center ("Date: " & Format(Now, "mmm d, yyyy"))
    Print_Line ""
    Print_Line ""
    Print_Line ""
    Print_Line ""
    'Print Header & _
    '"123456789^123456789^123456789^123456789^123456789^123456789^123456789^1234567890"
    For idx = 1 To NR_JOBSEQ
'        GetSequence MASTERMODE, idx, 0
'        If DspSequence.Description <> "                    " Then
'            Print_Line "Sequence Description : " & _
'                Format(idx, "#0            ") & _
'                Format(DspSequence.Description)
    '        Print_Line ("Working Capacity:                 (grams) " & _
    '            Format(DspSequence.WorkingCapacity, "###0.0"))
    '        Print_Line ("Working Capacity Volume:         (liters) " & _
    '            Format(DspSequence.WorkingVolume, "#0.00"))
'            Print_Line ""
'        End If
     Next idx
    
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
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub cmdSaveSeq_Click()
'    SaveJobSequence NewSeq
'End Sub
'
'Public Sub SaveJobSequence(iSeq As JobSequence)
'
' save Job Sequence
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 390, 1
    lblMessage.Caption = vbCrLf
    If iSeq.Validated Then
        JobSeqAutoEdit = True
        Select Case JobSeqMode
            Case MASTERMODE
                ' master
                ' Save Master Sequence Information
                SaveMasterSequence CInt(iSeq.Number)
                ' update displayed duration
                lblSeqDuration.Caption = iSeq.EstSeqDurDesc
                lblMessage.Caption = lblMessage.Caption & "Master JobSequence #" + Format(iSeq.Number, "###0") & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Saved to database" & vbCrLf
            Case STATIONMODE
                ' station
                StationSequence(DispStn, DispShift) = iSeq
                If StationSequence(DispStn, DispShift).Number <> DispSeqNum Then
                    StationSequence(DispStn, DispShift).Number = DispSeqNum
                End If
                ' save Station Sequence recipes
                Save_StationSequences
                Select Case NR_SHIFT
                    Case 1
                        lblMessage.Caption = lblMessage.Caption & "JobSequence Saved to Station #" + Format(DispStn, "0")
                    Case 2
                        lblMessage.Caption = lblMessage.Caption & "JobSequence Saved to Station #" + Format(DispStn, "0") + " / Shift #" + Format(DispShift, "0")
                End Select
                lblMessage.Caption = lblMessage.Caption & vbCrLf
        End Select
        ' update curr
        CurrSeq = iSeq
        ' update prev
        PrevSeq = iSeq
        ' update display
        DisplaySequence iSeq
        Reset_BackColors
        ' update control flags
        JobSeqAutoEdit = False
        bUnSavedSeq = False
        ' setup control buttons
        SetupForEdit False
        
    Else
        lblMessage.Caption = lblMessage.Caption & "JobSequence is Not Valid" & vbCrLf
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

Private Sub cmdRestoreSeq_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 390, 1111

    lblMessage.Caption = vbCrLf
    JobSeqAutoEdit = True
    CurrSeq = PrevSeq
    DisplaySequence CurrSeq
    JobSeqAutoEdit = False
    lblMessage.Caption = lblMessage.Caption & "JobSequence Restored" & vbCrLf
    bUnSavedSeq = False
    bUnValidSeq = False
    bUnUpdated = False
    ' setup control buttons
    SetupForEdit bEditing
    
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

Private Sub cmdDnRec_Click()
    If SelectedCourse < CourseCount Then SelectedCourse = SelectedCourse + 1
    DisplaySelectedCourse NewSeq
End Sub

Private Sub cmdUpRec_Click()
    If SelectedCourse > 1 Then SelectedCourse = SelectedCourse - 1
    DisplaySelectedCourse NewSeq
End Sub

Private Sub cmdNewCourse_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 390, 1121

    SelectedCourse = NewSeq.NumCourses + 1
    CourseCount = CInt(SelectedCourse)
    NewSeq.NumCourses = CInt(SelectedCourse)
    txtRecCourse.text = Format(SelectedCourse, "##0")
    txtRecType.text = Format(courseRecipe, "##0")
    txtRecPause.text = "0"
    txtRecRecipe.text = "0"
    txtRecCycles.text = "0"
    txtRecLoadRate.text = "0"
    txtRecPurgeRate.text = "0"
    txtRecMsgText.text = " "
    CopyScreenToCourse NewSeq
    ReSeqCourses NewSeq
    DisplaySequence NewSeq
    ' setup control buttons
    bUnUpdated = True
    SetupForEdit bEditing
    
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

Private Sub cmdDelete_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 390, 1131

    lblMessage.Caption = vbCrLf
    ReSeqCourses NewSeq
    If (NewSeq.NumCourses > 1) Then
        NewSeq.CourseData(SelectedCourse).Type = courseUndefined
        SelectedCourse = NewSeq.NumCourses - 1
        CourseCount = CInt(SelectedCourse)
        NewSeq.NumCourses = CInt(SelectedCourse)
        ReSeqCourses NewSeq
        If (NewSeq.CourseData(SelectedCourse).Type = courseUndefined) Then SelectedCourse = 1
        DisplaySequence NewSeq
        SetupForEdit bEditing
        lblMessage.Caption = lblMessage.Caption & "Course Deleted" & vbCrLf
    Else
        lblMessage.Caption = lblMessage.Caption & "Only 1 Course" & vbCrLf
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

Private Sub cmdUpdate_Click()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 390, 1140
Dim flag As Boolean

    flag = True
    lblMessage.Caption = vbCrLf
    If Not Range_Check(txtRecCourse, 1, 99, "Course Number") Then flag = False
    If Not Range_Check(txtRecType, courseWait, courseRecipe, "Course Type") Then flag = False
    If flag Then
        Select Case CInt(txtRecType.text)
            Case courseWait
                ChgErrModule 390, 1141
                If (Len(txtRecMsgText.text) < 3) Then
                    txtRecMsgText.BackColor = EntryInvalid_BackColor
                    lblMessage.Caption = lblMessage.Caption & "MessageBox Text" & vbCrLf
                End If
            Case coursePause
                ChgErrModule 390, 1142
                If Not Range_Check(txtRecPause, 1, 99999, "Pause Duration") Then flag = False
            Case courseRecipe
                ChgErrModule 390, 1143
                ' run recipe with optional changes
                If Not Range_Check(txtRecRecipe, 0, MAX_RCP, "Recipe Number") Then flag = False
                If Not Range_Check(txtRecCycles, 0, 999, "Cycles") Then flag = False
                If Not Range_Check(txtRecLoadRate, 0, 999, "Load Rate") Then flag = False
                If Not Range_Check(txtRecPurgeRate, 0, 999, "Purge Rate") Then flag = False
        End Select
    End If
    If flag Then
        ChgErrModule 390, 1149
        Reset_BackColors
        JobSeqAutoEdit = True
        CopyScreenToCourse NewSeq
        DisplaySequence NewSeq
        JobSeqAutoEdit = False
        lblMessage.Caption = lblMessage.Caption & "Course Updated" & vbCrLf
        bUnUpdated = False
        SeqHasChanged
    End If
    ' setup control buttons
    SetupForEdit bEditing

ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    lblMessage = "Update Failed" & vbCrLf
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Private Sub dbgCourses_AfterInsert()
    KeyPreview = True
End Sub

Private Sub dbgCourses_BeforeInsert(Cancel As Integer)
    KeyPreview = True
End Sub

Private Sub dbgCourses_Change()
    KeyPreview = True
End Sub

Private Sub dbgCourses_Click()
    KeyPreview = True
End Sub

Private Sub dbgCourses_ColEdit(ByVal ColIndex As Integer)
    KeyPreview = True
    Select Case JobSeqMode
        Case MASTERMODE
            ' master
            dbgCourses.Columns(0).Value = DispSeqNum
        Case STATIONMODE
            ' station
            dbgCourses.Columns(0).Value = DispStn
            dbgCourses.Columns(1).Value = DispShift
    End Select
End Sub

Private Sub dbgCourses_HeadClick(ByVal ColIndex As Integer)
    KeyPreview = True
End Sub

Private Sub Form_Load()
    KeyPreview = True
    bEmptyMemSeq = True
    bEditing = False
    bLoadCoursesOnly = False
    JobSeqAutoEdit = True
    frmSelCourse.ForeColor = Titles_ForeColor
    txtSeqDescription.ForeColor = Titles_ForeColor
    txtDispSeqNum.ForeColor = Titles_ForeColor
    txtPriScale.ForeColor = Titles_ForeColor
    txtAuxScale.ForeColor = Titles_ForeColor
    txtRecCourse.ForeColor = Titles_ForeColor
    txtRecType.ForeColor = Titles_ForeColor
    txtRecPause.ForeColor = Titles_ForeColor
    txtRecRecipe.ForeColor = Titles_ForeColor
    txtRecCycles.ForeColor = Titles_ForeColor
    txtRecLoadRate.ForeColor = Titles_ForeColor
    txtRecPurgeRate.ForeColor = Titles_ForeColor
    txtRecMsgText.ForeColor = Titles_ForeColor
    lblMessage.ForeColor = Message_ForeColor
    ' Line Volume
    pnlLineVolume.Top = OutOfSight
    If USINGLINEVOLUME Then
        cmdLineVolume.Visible = True
        If USINGLVol_SI Then
            ' USING SI UNITS
            lblIDLoad.Caption = "mm"
            lblIDPurge.Caption = "mm"
            lblIDVent.Caption = "mm"
            lblLoadL.Caption = "m"
            lblPurgeL.Caption = "m"
            lblVentL.Caption = "m"
            txtIDLoad.ToolTipText = "0 to 25.4 millimeters"
            txtIDPurge.ToolTipText = "0 to 25.4 millimeters"
            txtIDVent.ToolTipText = "0 to 25.4 millimeters"
            txtLoadL.ToolTipText = "0 to 60.96 meters"
            txtPurgeL.ToolTipText = "0 to 60.96 meters"
            txtVentL.ToolTipText = "0 to 60.96 meters"
        ElseIf USINGLVol_Engl Then
            ' USING ENGLISH UNITS
            lblIDLoad.Caption = "in"
            lblIDPurge.Caption = "in"
            lblIDVent.Caption = "in"
            lblLoadL.Caption = "ft"
            lblPurgeL.Caption = "ft"
            lblVentL.Caption = "ft"
            txtIDLoad.ToolTipText = "0 to 1 inches"
            txtIDPurge.ToolTipText = "0 to 1 inches"
            txtIDVent.ToolTipText = "0 to 1 inches"
            txtLoadL.ToolTipText = "0 to 200 feet"
            txtPurgeL.ToolTipText = "0 to 200 feet"
            txtVentL.ToolTipText = "0 to 200 feet"
        Else
            ' USING UNKNOWN UNITS
            lblIDLoad.Caption = "??"
            lblIDPurge.Caption = "??"
            lblIDVent.Caption = "??"
            lblLoadL.Caption = "??"
            lblPurgeL.Caption = "??"
            lblVentL.Caption = "??"
            lblIDLoad.ToolTipText = "LOAD Inside Diameters in ??"
            lblIDPurge.ToolTipText = "PURGE Inside Diameters in ??"
            lblIDVent.ToolTipText = "VENT Inside Diameters in ??"
            lblLoadL.ToolTipText = "LOAD Length in ??"
            lblPurgeL.ToolTipText = "PURGE Length in ??"
            lblVentL.ToolTipText = "VENT Length in ??"
        End If
    Else
        cmdLineVolume.Visible = False
    End If
    Dbg_EditHeight = 4155
    Dbg_FullHeight = 7155
    cmdOpen.Top = cmdCopy.Top
    cmdNewSeq.Top = cmdCopy.Top
    cmdLoadAll.Top = cmdCopy.Top
    cmdLoadCourses.Top = cmdCopy.Top
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
Exit Sub
'    Criteria = "SELECT * FROM [StationSequence] WHERE [Station] = " & DispStn & "  and [Shift] = " & DispShift & " "
'    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
'    Criteria = "SELECT * FROM [StationSequenceCourses] WHERE [Station] = " & DispStn & "  and [Shift] = " & DispShift & "ORDER BY [CourseNumber] DESC" & " "
'    frmCourses.adoCourses.RecordSource = Criteria
'    JobSeqMode = STATIONMODE
'    InitSeqLoad  ' open Sequence / recipe database
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ExitScreen
End Sub

Private Sub lblNumCourses_Click()
'    Select Case JobSeqMode
'        Case MASTERMODE
'            ' master
'            RefreshCourses MASTERMODE, DispSeqNum, 0
'        Case STATIONMODE
'            ' station
'            RefreshCourses STATIONMODE, DispStn, DispShift
'    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        ExitScreen
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Function Range_Check(tcontrol As Control, slow, shigh As Single, slabel As String) As Boolean
' Function Name:    Range_Check
' Description:      Checks the value of the control text entry and compares
'                   it to the low and high range limits provided.  If the
'                   value is outside of the range given, or if the entry
'                   is not a valid numeric entry, an error message is
'                   displayed.  The error message is preceeded by the
'                   label provided in slabel.
'
' tcontrol          control name whose text value will be checked
' slow              low range value, single
' shigh             high range value, single
' slabel            string containing label for error message,
'                   if slabel = "Date" message will be;
'                   Date: Value out of Range!
'
Dim svalue As Single
Dim Message As String

SetErrModule 390, 3
If UseLocalErrorHandler Then On Error GoTo localhandler

    Range_Check = True
    
    If (tcontrol.text = Empty) Then
        
        ' Empty Value
        Range_Check = False
        Message = slabel & ":  Value is Empty!"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    
    ElseIf Not IsNumeric(tcontrol.text) Then
        
        ' Non-Numeric Value
        Range_Check = False
        Message = slabel & ":  Value is Not Numeric!"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    
    Else
    
        ' Numeric Value
        svalue = CSng(tcontrol.text)
        
        ' Check Value against Limits
        If svalue < slow Or svalue > shigh Then
            Range_Check = False
            tcontrol.BackColor = EntryInvalid_BackColor
        '    tcontrol.SelStart = 0
        '    tcontrol.SelLength = Len(tcontrol.text)
        '    tcontrol.SetFocus
            Message = slabel & ":  Value out of range! " & "( " & Format(slow, "###0.00") & " - " & Format(shigh, "###0.00") & " )"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        End If
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

Private Sub Reset_BackColors()
'
' resets the background colors
'
SetErrModule 390, 6

If UseLocalErrorHandler Then On Error GoTo localhandler

    lblCourses.BackColor = txtNotHighlight.BackColor
    lblSeqDuration.BackColor = txtNotHighlight.BackColor
    txtSeqDescription.BackColor = txtNotHighlight.BackColor
    txtPriScale.BackColor = txtNotHighlight.BackColor
    txtAuxScale.BackColor = txtNotHighlight.BackColor
    txtRecCourse.BackColor = txtNotHighlight.BackColor
    txtRecType.BackColor = txtNotHighlight.BackColor
    txtRecPause.BackColor = txtNotHighlight.BackColor
    txtRecRecipe.BackColor = txtNotHighlight.BackColor
    txtRecCycles.BackColor = txtNotHighlight.BackColor
    txtRecLoadRate.BackColor = txtNotHighlight.BackColor
    txtRecPurgeRate.BackColor = txtNotHighlight.BackColor
    txtRecMsgText.BackColor = txtNotHighlight.BackColor

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

Private Sub txtDispSeqNum_Click()
'    frmSearchSeq.SetInitialRow (DispSeqNum)
'    frmSearchSeq.Show
End Sub

Private Sub txtRecCourse_Change()
'    CourseDataHasChanged
End Sub

Private Sub txtRecCycles_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecLoadRate_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecMsgText_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecPause_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecPurgeRate_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecRecipe_Change()
    CourseDataHasChanged
End Sub

Private Sub txtRecType_Change()
    CourseDataHasChanged
End Sub

Private Sub txtSeqDescription_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
        Case False
            txtSeqDescription.text = CurrSeq.Description
    End Select
End Sub

Private Sub txtAuxScale_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtAuxScale.BackColor = txtNotHighlight.BackColor
        Case False
            txtAuxScale.text = Format(CurrSeq.AuxScaleNo, "#0")
    End Select
End Sub

Private Sub txtPriScale_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtPriScale.BackColor = txtNotHighlight.BackColor
        Case False
            txtPriScale.text = Format(CurrSeq.PriScaleNo, "#0")
    End Select
End Sub

Private Sub txtIDLoad_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtIDLoad.BackColor = txtNotHighlight.BackColor
        Case False
            txtIDLoad.text = Format(CurrSeq.IDLoad, "#0.00")
    End Select
End Sub

Private Sub txtIDPurge_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtIDPurge.BackColor = txtNotHighlight.BackColor
        Case False
            txtIDPurge.text = Format(CurrSeq.IDPurge, "#0.00")
    End Select
End Sub

Private Sub txtIDVent_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtIDVent.BackColor = txtNotHighlight.BackColor
        Case False
            txtIDVent.text = Format(CurrSeq.IDVent, "#0.00")
    End Select
End Sub

Private Sub txtLoadL_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtLoadL.BackColor = txtNotHighlight.BackColor
        Case False
            txtLoadL.text = Format(CurrSeq.LoadL, "##0.00")
    End Select
End Sub

Private Sub txtLoadV_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtLoadV.BackColor = txtNotHighlight.BackColor
        Case False
            txtLoadV.text = Format(CurrSeq.LoadV, "##0.00")
    End Select
End Sub

Private Sub txtPurgeL_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtPurgeL.BackColor = txtNotHighlight.BackColor
        Case False
            txtPurgeL.text = Format(CurrSeq.PurgeL, "##0.00")
    End Select
End Sub

Private Sub txtPurgeV_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
            txtPurgeV.BackColor = txtNotHighlight.BackColor
        Case False
            txtPurgeV.text = Format(CurrSeq.PurgeV, "##0.00")
    End Select
End Sub

Private Sub txtVentL_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
'            lblMessage.Caption = vbCrLf
            txtVentL.BackColor = txtNotHighlight.BackColor
        Case False
            txtVentL.text = Format(CurrSeq.VentL, "##0.00")
    End Select
End Sub

Private Sub txtVentV_Change()
    Select Case bEditing
        Case True
            SeqHasChanged
'            lblMessage.Caption = vbCrLf
            txtVentV.BackColor = txtNotHighlight.BackColor
        Case False
            txtVentV.text = Format(CurrSeq.VentV, "##0.00")
    End Select
End Sub

Private Sub ClearCoursesDataTable()
'
' Clear zCourses DataTable
'
On Error GoTo localhandler
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
    ' Clear Courses DataTable
    rsCrit = "SELECT * FROM [zCourses]"
    Set dB = OpenDatabase(FILEPATH_rcp & DATARCP)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    If Not rS.BOF Then
        rS.MoveLast
        Do While Not rS.BOF
            rS.Delete
            rS.MoveLast
        Loop
    End If
    rS.Close
    dB.Close
Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Clearing zCourses DataTable: " & error$(err)
    etxt = Mid$(etxt, 1, 255)
    Write_ELog etxt
    Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Private Sub ClearCourses(iSeq As JobSequence)
Dim iCourse As Integer
    For iCourse = 1 To MAX_COURSES
        iSeq.CourseData(iCourse).CourseNumber = iCourse
        iSeq.CourseData(iCourse).Type = courseUndefined
        iSeq.CourseData(iCourse).EstCourseDuration = CSng(0)
        iSeq.CourseData(iCourse).RecipeNumber = CInt(0)
        iSeq.CourseData(iCourse).Cycles = CInt(0)
        iSeq.CourseData(iCourse).LoadRate = CSng(0)
        iSeq.CourseData(iCourse).MsgText = "none"
        iSeq.CourseData(iCourse).PauseDuration = CSng(0)
        iSeq.CourseData(iCourse).PurgeRate = CSng(0)
    Next iCourse
End Sub

Private Sub CopyStationToCurr(ByVal iStation As Integer, ByVal iShift As Integer)
'
' Copy data from the Station JobSequence to the Current JobSequence
'
    ' Copy Station JobSequence to Curr JobSequence
    CurrSeq = StationSequence(DispStn, DispShift)
    ' set the number of courses (rows)
    CurrSeq.NumCourses = GetCourseCount(CurrSeq)
    CourseCount = CurrSeq.NumCourses
End Sub

Private Sub CopyMasterToCurr(ByVal iMaster As Integer)
'
' Copy data from the Master JobSequence to the Current JobSequence
'
Dim dB As Database
Dim src_rS As Recordset
Dim src_rsCrit As String
Dim iCourse As Integer
    
    ' Clear Current Courses
    ClearCourses CurrSeq
    ' Open Database for copying
    Set dB = OpenDatabase(FILEPATH_rcp & DATARCP)

    ' Open the Master JobSequence Table
    src_rsCrit = "SELECT * FROM [MasterSequence] "
    src_rsCrit = src_rsCrit & "WHERE [Number] = " & iMaster & " "
    Set src_rS = dB.OpenRecordset(src_rsCrit, dbOpenDynaset)
    If Not src_rS.BOF Then
        ' copy the values
        src_rS.MoveFirst
        CurrSeq.Number = src_rS("Number")
        CurrSeq.Description = src_rS("Description")
        CurrSeq.NumCourses = src_rS("Courses")
        CurrSeq.PriScaleNo = src_rS("PriScale")
        CurrSeq.AuxScaleNo = src_rS("AuxScale")
        CurrSeq.IDLoad = src_rS("IDLoad")
        CurrSeq.IDPurge = src_rS("IDPurge")
        CurrSeq.IDVent = src_rS("IDVent")
        CurrSeq.LoadL = src_rS("LoadL")
        CurrSeq.LoadV = src_rS("LoadV")
        CurrSeq.PurgeL = src_rS("PurgeL")
        CurrSeq.PurgeV = src_rS("PurgeV")
        CurrSeq.VentL = src_rS("VentL")
        CurrSeq.VentV = src_rS("VentV")
        CurrSeq.Validated = src_rS("Validated")
        CurrSeq.EstSeqDuration = src_rS("EstSeqDuration")
        CurrSeq.EstSeqDurDesc = src_rS("EstSeqDurDesc")
    Else
        ' clear the values
        CurrSeq.Number = iMaster
        CurrSeq.Description = "unnamed"
        CurrSeq.NumCourses = 0
        CurrSeq.PriScaleNo = 0
        CurrSeq.AuxScaleNo = 0
        CurrSeq.IDLoad = 0
        CurrSeq.IDPurge = 0
        CurrSeq.IDVent = 0
        CurrSeq.LoadL = 0
        CurrSeq.LoadV = 0
        CurrSeq.PurgeL = 0
        CurrSeq.PurgeV = 0
        CurrSeq.VentL = 0
        CurrSeq.VentV = 0
        CurrSeq.Validated = False
        CurrSeq.EstSeqDuration = 0
        CurrSeq.EstSeqDurDesc = "undefined"
    End If
    src_rS.Close
    
    ' Open the Master JobSequence Courses Table
    src_rsCrit = "SELECT * FROM [MasterSequenceCourses] "
    src_rsCrit = src_rsCrit & "WHERE [SeqNum] = " & iMaster & " "
    src_rsCrit = src_rsCrit & " ORDER BY [MasterSequenceCourses].[CourseNumber] ASC "
    Set src_rS = dB.OpenRecordset(src_rsCrit, dbOpenDynaset)
    If Not src_rS.BOF Then
        src_rS.MoveLast
        src_rS.MoveFirst
        ' set the number of courses (rows)
        CourseCount = src_rS.RecordCount
        ' Copy records
        Do While Not src_rS.EOF
            iCourse = src_rS("CourseNumber")
            CurrSeq.CourseData(iCourse).CourseNumber = iCourse
            CurrSeq.CourseData(iCourse).Type = src_rS("Type")
            CurrSeq.CourseData(iCourse).OkToProceed = False
            CurrSeq.CourseData(iCourse).PauseDuration = src_rS("PauseDuration")
            CurrSeq.CourseData(iCourse).RecipeNumber = src_rS("RecipeNumber")
            CurrSeq.CourseData(iCourse).Cycles = src_rS("Cycles")
            CurrSeq.CourseData(iCourse).LoadRate = src_rS("LoadRate")
            CurrSeq.CourseData(iCourse).PurgeRate = src_rS("PurgeRate")
            CurrSeq.CourseData(iCourse).EstCourseDuration = src_rS("EstCourseDuration")
            If (IsNull(src_rS("MsgText"))) Then
                CurrSeq.CourseData(iCourse).MsgText = "none"
            Else
                CurrSeq.CourseData(iCourse).MsgText = src_rS("MsgText")
            End If
            CurrSeq.CourseData(iCourse).EstCourseDuration = src_rS("EstCourseDuration")
            src_rS.MoveNext
        Loop
    End If
    
    'close source recordset
    src_rS.Close
    'close database
    dB.Close
    
    ' set the number of courses (rows)
    CurrSeq.NumCourses = GetCourseCount(CurrSeq)
    CourseCount = CurrSeq.NumCourses

End Sub

Private Sub CopyMasterToNew(ByVal iMaster As Integer)
'
' Copy data from the Master JobSequence to the New JobSequence
'
Dim dB As Database
Dim src_rS As Recordset
Dim src_rsCrit As String
Dim iCourse As Integer
    
    ' Clear New Courses
    ClearCourses NewSeq
    ' Open Database for copying
    Set dB = OpenDatabase(FILEPATH_rcp & DATARCP)

    If (Not bLoadCoursesOnly) Then
        ' Open the Master JobSequence Table
        src_rsCrit = "SELECT * FROM [MasterSequence] "
        src_rsCrit = src_rsCrit & "WHERE [Number] = " & iMaster & " "
        Set src_rS = dB.OpenRecordset(src_rsCrit, dbOpenDynaset)
        If Not src_rS.BOF Then
            ' copy the values
            src_rS.MoveFirst
            NewSeq.Number = src_rS("Number")
            NewSeq.Description = src_rS("Description")
            NewSeq.NumCourses = src_rS("Courses")
            NewSeq.PriScaleNo = src_rS("PriScale")
            NewSeq.AuxScaleNo = src_rS("AuxScale")
            NewSeq.IDLoad = src_rS("IDLoad")
            NewSeq.IDPurge = src_rS("IDPurge")
            NewSeq.IDVent = src_rS("IDVent")
            NewSeq.LoadL = src_rS("LoadL")
            NewSeq.LoadV = src_rS("LoadV")
            NewSeq.PurgeL = src_rS("PurgeL")
            NewSeq.PurgeV = src_rS("PurgeV")
            NewSeq.VentL = src_rS("VentL")
            NewSeq.VentV = src_rS("VentV")
            NewSeq.Validated = src_rS("Validated")
            NewSeq.EstSeqDuration = src_rS("EstSeqDuration")
            NewSeq.EstSeqDurDesc = src_rS("EstSeqDurDesc")
        Else
            ' clear the values
            NewSeq.Number = iMaster
            NewSeq.Description = "unnamed"
            NewSeq.NumCourses = 0
            NewSeq.PriScaleNo = 0
            NewSeq.AuxScaleNo = 0
            NewSeq.IDLoad = 0
            NewSeq.IDPurge = 0
            NewSeq.IDVent = 0
            NewSeq.LoadL = 0
            NewSeq.LoadV = 0
            NewSeq.PurgeL = 0
            NewSeq.PurgeV = 0
            NewSeq.VentL = 0
            NewSeq.VentV = 0
            NewSeq.Validated = False
            NewSeq.EstSeqDuration = 0
            NewSeq.EstSeqDurDesc = "undefined"
        End If
        src_rS.Close
    End If
    
    ' Open the Master JobSequence Courses Table
    src_rsCrit = "SELECT * FROM [MasterSequenceCourses] "
    src_rsCrit = src_rsCrit & "WHERE [SeqNum] = " & iMaster & " "
    src_rsCrit = src_rsCrit & " ORDER BY [MasterSequenceCourses].[CourseNumber] ASC "
    Set src_rS = dB.OpenRecordset(src_rsCrit, dbOpenDynaset)
    If Not src_rS.BOF Then
        src_rS.MoveLast
        src_rS.MoveFirst
        ' Copy records
        Do While Not src_rS.EOF
            iCourse = src_rS("CourseNumber")
            NewSeq.CourseData(iCourse).CourseNumber = iCourse
            NewSeq.CourseData(iCourse).Type = src_rS("Type")
            NewSeq.CourseData(iCourse).OkToProceed = False
            NewSeq.CourseData(iCourse).PauseDuration = src_rS("PauseDuration")
            NewSeq.CourseData(iCourse).RecipeNumber = src_rS("RecipeNumber")
            NewSeq.CourseData(iCourse).Cycles = src_rS("Cycles")
            NewSeq.CourseData(iCourse).LoadRate = src_rS("LoadRate")
            NewSeq.CourseData(iCourse).PurgeRate = src_rS("PurgeRate")
            NewSeq.CourseData(iCourse).EstCourseDuration = src_rS("EstCourseDuration")
            If (IsNull(src_rS("MsgText"))) Then
                NewSeq.CourseData(iCourse).MsgText = "none"
            Else
                NewSeq.CourseData(iCourse).MsgText = src_rS("MsgText")
            End If
            NewSeq.CourseData(iCourse).EstCourseDuration = src_rS("EstCourseDuration")
            src_rS.MoveNext
        Loop
    End If
    
    'close source recordset
    src_rS.Close
    'close database
    dB.Close
    
    ' check for out-of-sequence courses
    ReSeqCourses NewSeq
    ' remember the number of courses
    NewSeq.NumCourses = GetCourseCount(NewSeq)
    CourseCount = NewSeq.NumCourses
    ' reset course index
    SelectedCourse = 1
    
End Sub

Private Sub DisplaySelectedCourse(iSeq As JobSequence)
'
    txtRecCourse.text = Format(iSeq.CourseData(SelectedCourse).CourseNumber, "##0")
    txtRecType.text = Format(iSeq.CourseData(SelectedCourse).Type, "#0")
    txtRecPause.text = Format(iSeq.CourseData(SelectedCourse).PauseDuration, "##0.00")
    txtRecRecipe.text = Format(iSeq.CourseData(SelectedCourse).RecipeNumber, "##0")
    txtRecCycles.text = Format(iSeq.CourseData(SelectedCourse).Cycles, "#,##0")
    txtRecLoadRate.text = Format(iSeq.CourseData(SelectedCourse).LoadRate, "##0.00")
    txtRecPurgeRate.text = Format(iSeq.CourseData(SelectedCourse).PurgeRate, "##0.00")
    txtRecMsgText.text = iSeq.CourseData(SelectedCourse).MsgText
End Sub

Private Sub DisplaySequence(iSeq As JobSequence)
'
On Error GoTo localhandler
Dim dB As Database
Dim rS As Recordset
Dim rsCrit As String
Dim iCourse As Integer

    ' update screen jobseq values
    txtDispSeqNum.Caption = Format(iSeq.Number, "#00")
    txtSeqDescription.text = iSeq.Description
    lblSeqDuration.Caption = iSeq.EstSeqDurDesc
    lblCourses.Caption = Format(iSeq.NumCourses, "####0")
    txtPriScale.text = Format(iSeq.PriScaleNo, "####0")
    txtAuxScale.text = Format(iSeq.AuxScaleNo, "####0")
    txtIDLoad.text = Format(iSeq.IDLoad, "#0.00")
    txtIDPurge.text = Format(iSeq.IDPurge, "#0.00")
    txtIDVent.text = Format(iSeq.IDVent, "#0.00")
    txtLoadL.text = Format(iSeq.LoadL, "##0.00")
    txtLoadV.text = Format(iSeq.LoadV, "##0.00")
    txtPurgeL.text = Format(iSeq.PurgeL, "##0.00")
    txtPurgeV.text = Format(iSeq.PurgeV, "##0.00")
    txtVentL.text = Format(iSeq.VentL, "##0.00")
    txtVentV.text = Format(iSeq.VentV, "##0.00")
    
    ' Clear zCourses DataTable
    ClearCoursesDataTable
    ' Update zCourses DataTable
    rsCrit = "SELECT * FROM [zCourses]"
    Set dB = OpenDatabase(FILEPATH_rcp & DATARCP)
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)
    For iCourse = 1 To MAX_COURSES
        If (iSeq.CourseData(iCourse).Type <> courseUndefined) Then
            rS.AddNew
                rS("CourseNumber") = iSeq.CourseData(iCourse).CourseNumber
                rS("Type") = iSeq.CourseData(iCourse).Type
                rS("PauseDuration") = iSeq.CourseData(iCourse).PauseDuration
                rS("RecipeNumber") = iSeq.CourseData(iCourse).RecipeNumber
                rS("Cycles") = iSeq.CourseData(iCourse).Cycles
                rS("LoadRate") = iSeq.CourseData(iCourse).LoadRate
                rS("PurgeRate") = iSeq.CourseData(iCourse).PurgeRate
                rS("EstCourseDuration") = iSeq.CourseData(iCourse).EstCourseDuration
                rS("MsgText") = iSeq.CourseData(iCourse).MsgText
            rS.Update
        End If
    Next iCourse
    ' close DataTable and DB
    rS.Close
    dB.Close
    
    ' refresh Courses Display
    adoCourses.RecordSource = "SELECT * FROM [zCourses]  ORDER BY [zCourses].[CourseNumber] ASC "
    adoCourses.Refresh

    Reset_BackColors

Exit Sub

localhandler:
Dim iresponse As Integer
Dim etxt As String
    etxt = "Error " & err & " Updating zCourses DataTable: " & error$(err)
    etxt = Mid$(etxt, 1, 255)
    Write_ELog etxt
    Delay_Box etxt, MSGDELAY, msgSHOW
End Sub

Private Sub CopyCourses(srcSeq As JobSequence, dstSeq As JobSequence)
Dim iCourse As Integer
    For iCourse = 1 To MAX_COURSES
        dstSeq.CourseData(iCourse).CourseNumber = srcSeq.CourseData(iCourse).CourseNumber
        dstSeq.CourseData(iCourse).Type = srcSeq.CourseData(iCourse).Type
        dstSeq.CourseData(iCourse).OkToProceed = srcSeq.CourseData(iCourse).OkToProceed
        dstSeq.CourseData(iCourse).PauseDuration = srcSeq.CourseData(iCourse).PauseDuration
        dstSeq.CourseData(iCourse).RecipeNumber = srcSeq.CourseData(iCourse).RecipeNumber
        dstSeq.CourseData(iCourse).Cycles = srcSeq.CourseData(iCourse).Cycles
        dstSeq.CourseData(iCourse).LoadRate = srcSeq.CourseData(iCourse).LoadRate
        dstSeq.CourseData(iCourse).PurgeRate = srcSeq.CourseData(iCourse).PurgeRate
        dstSeq.CourseData(iCourse).EstCourseDuration = srcSeq.CourseData(iCourse).EstCourseDuration
        dstSeq.CourseData(iCourse).MsgText = srcSeq.CourseData(iCourse).MsgText
    Next iCourse
End Sub

Private Sub CopyScreenToSeq(iSeq As JobSequence)
    iSeq.Number = CInt(txtDispSeqNum.Caption)
    iSeq.Description = txtSeqDescription.text
    iSeq.NumCourses = CSng(lblCourses.Caption)
    iSeq.PriScaleNo = CSng(txtPriScale.text)
    iSeq.AuxScaleNo = CSng(txtAuxScale.text)
    iSeq.IDLoad = CSng(txtIDLoad.text)
    iSeq.IDPurge = CSng(txtIDPurge.text)
    iSeq.IDVent = CSng(txtIDVent.text)
    iSeq.LoadL = CSng(txtLoadL.text)
    iSeq.LoadV = CSng(txtLoadV.text)
    iSeq.PurgeL = CSng(txtPurgeL.text)
    iSeq.PurgeV = CSng(txtPurgeV.text)
    iSeq.VentL = CSng(txtVentL.text)
    iSeq.VentV = CSng(txtVentV.text)
End Sub

Private Sub CopyScreenToCourse(iSeq As JobSequence)
'
    iSeq.CourseData(SelectedCourse).CourseNumber = CInt(txtRecCourse.text)
    iSeq.CourseData(SelectedCourse).Type = CInt(txtRecType.text)
    iSeq.CourseData(SelectedCourse).OkToProceed = False
    iSeq.CourseData(SelectedCourse).PauseDuration = CSng(txtRecPause.text)
    iSeq.CourseData(SelectedCourse).RecipeNumber = CInt(txtRecRecipe.text)
    iSeq.CourseData(SelectedCourse).Cycles = CInt(txtRecCycles.text)
    iSeq.CourseData(SelectedCourse).LoadRate = CSng(txtRecLoadRate.text)
    iSeq.CourseData(SelectedCourse).PurgeRate = CSng(txtRecPurgeRate.text)
    iSeq.CourseData(SelectedCourse).EstCourseDuration = CSng(0)
    iSeq.CourseData(SelectedCourse).MsgText = txtRecMsgText.text

End Sub


Private Sub SetButtonsForMode(ByVal MstStnMode As Integer)
    Select Case MstStnMode
        Case MASTERMODE
            cmdOpen.Visible = True
        '    cmdCopy.Enabled = False
        '    cmdPaste.Enabled = False
            cmdNewSeq.Visible = False
            cmdLoadAll.Visible = False
            cmdLoadCourses.Visible = False
            cmdPgDn.Visible = True
            cmdDn.Visible = True
            txtDispSeqNum.Visible = True
            cmdUp.Visible = True
            cmdPgUp.Visible = True
        '    cmdLoadDefault.Enabled = False
        '    cmdSaveSeq.Enabled = False
        '    cmdValidateSeq.Enabled = False
        '    cmdEditSeq.Enabled = True
        '    cmdRestoreSeq.Enabled = False
        '    cmdNewCourse.Enabled = False
        '    cmdUpdate.Enabled = False
        '    cmdDelete.Enabled = False
        Case STATIONMODE
            cmdOpen.Visible = False
        '    cmdCopy.Enabled = False
        '    cmdPaste.Enabled = False
            cmdNewSeq.Visible = True
            cmdLoadAll.Visible = True
            cmdLoadCourses.Visible = True
            cmdPgDn.Visible = False
            cmdDn.Visible = False
            txtDispSeqNum.Visible = False
            cmdUp.Visible = False
            cmdPgUp.Visible = False
        '    cmdLoadDefault.Enabled = False
        '    cmdSaveSeq.Enabled = False
        '    cmdValidateSeq.Enabled = False
        '    cmdEditSeq.Enabled = True
        '    cmdRestoreSeq.Enabled = False
        '    cmdNewCourse.Enabled = False
        '    cmdUpdate.Enabled = False
        '    cmdDelete.Enabled = False
    End Select
End Sub

Private Sub SetupForEdit(ByVal editFlag As Boolean)
    Select Case editFlag
        Case True
            ' EDITING
            bEditing = True
            cmdCopy.Enabled = True
            cmdPaste.Enabled = IIf(bEmptyMemSeq, False, True)
            cmdNewSeq.Enabled = True
            cmdLoadAll.Enabled = True
            cmdLoadCourses.Enabled = True
            cmdLoadDefault.Enabled = True
            cmdSaveSeq.Enabled = IIf(bUnSavedSeq, True, False)
            cmdValidateSeq.Enabled = IIf(bUnValidSeq, True, False)
            cmdEditSeq.Enabled = True
            cmdEditSeq.Caption = "Cancel Edit"
            cmdEditSeq.ToolTipText = "Discard editing changes"
            cmdRestoreSeq.Enabled = IIf(bUnSavedSeq, True, False)
            cmdNewCourse.Enabled = True
            cmdUpdate.Enabled = IIf(bUnUpdated, True, False)
            cmdDelete.Enabled = IIf((adoCourses.Recordset.RecordCount > 1), True, False)
            frmSelCourse.Left = dbgCourses.Left
            dbgCourses.Height = Dbg_EditHeight
        Case False
            ' VIEWING
            bEditing = False
            cmdCopy.Enabled = False
            cmdPaste.Enabled = False
            cmdNewSeq.Enabled = False
            cmdLoadAll.Enabled = False
            cmdLoadCourses.Enabled = False
            cmdLoadDefault.Enabled = False
            cmdSaveSeq.Enabled = IIf(bUnSavedSeq, True, False)
            cmdValidateSeq.Enabled = False
            cmdEditSeq.Enabled = True
            cmdEditSeq.Caption = " Edit Sequence"
            cmdEditSeq.ToolTipText = "Edit the JobSequence"
            cmdRestoreSeq.Enabled = False
            cmdNewCourse.Enabled = False
            cmdUpdate.Enabled = False
            cmdDelete.Enabled = False
            frmSelCourse.Left = OutOfSight
            dbgCourses.Height = Dbg_FullHeight
    End Select
End Sub

Private Sub ReSeqCourses(iSeq As JobSequence)
'
' Resequence courses; Eliminate gaps in course numbers
'
Dim idx As Integer
Dim idx2 As Integer
Dim flag As Boolean

    ' reset found-a-course-out-of-sequence flag
    flag = False
    ' copy source Seq to destination Seq
    tempSeq = iSeq
    ' clear destination Seq course data
    ClearCourses tempSeq
    ' init course index for destination Seq
    idx2 = 1
    ' for all the source courses
    For idx = 1 To MAX_COURSES
        ' source cource a defined course ??
        If (iSeq.CourseData(idx).Type <> courseUndefined) Then
            ' copy the source course data
            tempSeq.CourseData(idx2) = iSeq.CourseData(idx)
            ' set the destination course number
            tempSeq.CourseData(idx2).CourseNumber = idx2
            ' found-a-course-out-of-sequence ??
            If (idx <> idx2) Then flag = True
            ' increment destination course number
            idx2 = idx2 + 1
        End If
    Next idx
    ' inform user if courses were requenced
    If flag Then lblMessage.Caption = lblMessage.Caption & "Courses Resequenced" & vbCrLf
    ' update the CourseCount
    tempSeq.NumCourses = GetCourseCount(tempSeq)
    ' copy destination Seq back to source Seq
    iSeq = tempSeq

End Sub

Private Function GetCourseCount(iSeq As JobSequence) As Integer
Dim crsCount As Integer
Dim iCourse As Integer
    ' clear the counter
    crsCount = 0
    ' count the defined courses
    For iCourse = 1 To MAX_COURSES
        If iSeq.CourseData(iCourse).Type <> courseUndefined Then crsCount = crsCount + 1
    Next iCourse
    ' copy the count
    GetCourseCount = crsCount
End Function
