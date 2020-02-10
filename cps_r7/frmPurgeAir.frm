VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmPurgeAir 
   Caption         =   "PurgeAir Sources"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   630
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmSelector 
      Caption         =   "Selection"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   1215
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   14535
      Begin VB.CommandButton cmdUp 
         DisabledPicture =   "frmPurgeAir.frx":0000
         DownPicture     =   "frmPurgeAir.frx":0702
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
         Left            =   8970
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeAir.frx":0E04
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDown 
         DisabledPicture =   "frmPurgeAir.frx":1A46
         DownPicture     =   "frmPurgeAir.frx":2148
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
         Left            =   5520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeAir.frx":284A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin Threed.SSPanel pnlPurge 
         Height          =   840
         Left            =   6360
         TabIndex        =   40
         ToolTipText     =   "PurgeAir Source Number Displayed"
         Top             =   240
         Width           =   2610
         _Version        =   65536
         _ExtentX        =   4595
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "PurgeAir Source 9"
         ForeColor       =   -2147483646
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
   End
   Begin VB.Frame frmPurge 
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   6735
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   14535
      Begin VB.Frame frmPrgFunc 
         Caption         =   "PurgeAir Function Definition"
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
         Height          =   2775
         Left            =   120
         TabIndex        =   11
         Top             =   3840
         Width           =   14295
         Begin VB.Frame frmStnAnalog 
            Caption         =   "Analog Functions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   1020
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   14025
            Begin VB.CommandButton cmdaUp 
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
               Left            =   10680
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":348C
               Style           =   1  'Graphical
               TabIndex        =   42
               ToolTipText     =   "next"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.CommandButton cmdaDn 
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
               Left            =   11280
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":3B8E
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "previous"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.TextBox txtaFuncDesc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   48
               TabIndex        =   29
               Text            =   "Function Description123456789012345678901234567890"
               ToolTipText     =   "Function Description"
               Top             =   540
               Width           =   4125
            End
            Begin VB.TextBox txtaChan 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9840
               MaxLength       =   4
               TabIndex        =   28
               Text            =   "0"
               ToolTipText     =   "Opto Channel (0-15)"
               Top             =   540
               Width           =   600
            End
            Begin VB.TextBox txtaAddr 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9180
               MaxLength       =   2
               TabIndex        =   27
               Text            =   "1"
               ToolTipText     =   "Opto Address (0-49)"
               Top             =   540
               Width           =   600
            End
            Begin VB.TextBox txtaVDCMin 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7905
               MaxLength       =   6
               TabIndex        =   26
               Text            =   "0"
               ToolTipText     =   "Min Value in Volts"
               Top             =   540
               Width           =   1005
            End
            Begin VB.TextBox txtaVDCMax 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6840
               MaxLength       =   6
               TabIndex        =   25
               Text            =   "12345678"
               ToolTipText     =   "Max Value in Volts"
               Top             =   540
               Width           =   1005
            End
            Begin VB.TextBox txtaEUMin 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5565
               MaxLength       =   6
               TabIndex        =   24
               Text            =   "01234"
               ToolTipText     =   "Min Value in Engineering Units"
               Top             =   540
               Width           =   1005
            End
            Begin VB.TextBox txtaEUMax 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   4500
               MaxLength       =   6
               TabIndex        =   23
               Text            =   "12345"
               ToolTipText     =   "Max Value in Engineering Units"
               Top             =   540
               Width           =   1005
            End
            Begin VB.CommandButton cmdaSave 
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   12120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":4290
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Save Function Definition"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   1500
            End
            Begin VB.Label lblaFuncDesc 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
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
               TabIndex        =   36
               Top             =   300
               Width           =   4005
            End
            Begin VB.Label lblaChan 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Chan"
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
               Left            =   9840
               TabIndex        =   35
               Top             =   300
               Width           =   555
            End
            Begin VB.Label lblaAddr 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Addr"
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
               Left            =   9240
               TabIndex        =   34
               Top             =   300
               Width           =   555
            End
            Begin VB.Label lblaVdcMin 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Vdc Min"
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
               Left            =   7995
               TabIndex        =   33
               Top             =   300
               Width           =   885
            End
            Begin VB.Label lblaVdcMax 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Vdc Max"
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
               Left            =   6900
               TabIndex        =   32
               Top             =   300
               Width           =   915
            End
            Begin VB.Label lblaEUMin 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EU Min"
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
               Left            =   5655
               TabIndex        =   31
               Top             =   300
               Width           =   885
            End
            Begin VB.Label lblaEUMax 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "EU Max"
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
               Left            =   4560
               TabIndex        =   30
               Top             =   300
               Width           =   915
            End
         End
         Begin VB.Frame frmStnDigital 
            Caption         =   "Digital Functions"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   1020
            Left            =   120
            TabIndex        =   12
            Top             =   1620
            Width           =   14025
            Begin VB.CommandButton cmddUp 
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
               Left            =   10680
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":4992
               Style           =   1  'Graphical
               TabIndex        =   44
               ToolTipText     =   "next"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.CommandButton cmddDn 
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
               Left            =   11280
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":5094
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "previous"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   600
            End
            Begin VB.CheckBox chkdInverse 
               Alignment       =   1  'Right Justify
               Caption         =   "Use Inverse"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   7500
               TabIndex        =   17
               Top             =   540
               Width           =   1335
            End
            Begin VB.TextBox txtdFuncDesc 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               MaxLength       =   48
               TabIndex        =   16
               Text            =   "Function Description123456789012345678901234567890"
               ToolTipText     =   "Function Description"
               Top             =   540
               Width           =   4125
            End
            Begin VB.TextBox txtdAddr 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9180
               MaxLength       =   2
               TabIndex        =   15
               Text            =   "1"
               ToolTipText     =   "Opto Address (0-49)"
               Top             =   540
               Width           =   600
            End
            Begin VB.TextBox txtdChan 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   9840
               MaxLength       =   4
               TabIndex        =   14
               Text            =   "0"
               ToolTipText     =   "Opto Channel (0-15)"
               Top             =   540
               Width           =   600
            End
            Begin VB.CommandButton cmddSave 
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   480
               Left            =   12120
               MaskColor       =   &H00FFFFFF&
               Picture         =   "frmPurgeAir.frx":5796
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Save Function Definition"
               Top             =   360
               UseMaskColor    =   -1  'True
               Width           =   1500
            End
            Begin VB.Label lbldAddr 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Addr"
               Height          =   255
               Left            =   9240
               TabIndex        =   20
               Top             =   300
               Width           =   555
            End
            Begin VB.Label lbldChan 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Chan"
               Height          =   255
               Left            =   9840
               TabIndex        =   19
               Top             =   300
               Width           =   555
            End
            Begin VB.Label lbldFuncDesc 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   300
               Width           =   4005
            End
         End
      End
      Begin VB.Frame frmPurgeDef 
         Caption         =   "PurgeAir Source Definition"
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
         Height          =   3135
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   14295
         Begin VB.CheckBox chkAkPurgeRequest 
            Caption         =   " Using AK Purge Request?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   " Using Purge Request?"
            Top             =   720
            Width           =   3200
         End
         Begin VB.CheckBox chkAuxAir 
            Caption         =   " Using Aux Air Sol?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   " Using Vacuum Switch on Aspirator"
            Top             =   1455
            Width           =   3200
         End
         Begin VB.TextBox txtDescription 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6120
            MaxLength       =   60
            TabIndex        =   7
            ToolTipText     =   "Station Description (alphnumeric)"
            Top             =   360
            Width           =   3855
         End
         Begin VB.CheckBox chkVacSw 
            Caption         =   " Using Vacuum Switch?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   " Using Vacuum Switch on Aspirator"
            Top             =   1215
            Width           =   3200
         End
         Begin VB.CheckBox chkHdwPurgeRequest 
            Caption         =   " Using Hdw Purge Request?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   " Using Purge Request?"
            Top             =   975
            Width           =   3200
         End
         Begin VB.CheckBox chkPosPrsPrg 
            Caption         =   " Positive Pressure Purge Valves?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   " This Purge Air Source has Hardware for Positive Pressure Purgeing."
            Top             =   1680
            Width           =   3200
         End
         Begin VB.CommandButton cmdSave 
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
            Left            =   12600
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeAir.frx":5E98
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Save Information to File"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1500
         End
         Begin Threed.SSCommand cmdSetDefaults 
            Height          =   600
            Left            =   12600
            TabIndex        =   9
            ToolTipText     =   "Load & Save Default Values"
            Top             =   2280
            Visible         =   0   'False
            Width           =   1500
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1058
            _StockProps     =   78
            Caption         =   "&Defaults"
            ForeColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelWidth      =   4
         End
         Begin VB.Label lblDescription 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4800
            TabIndex        =   10
            Top             =   375
            Width           =   1080
         End
      End
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
      Left            =   13155
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmPurgeAir.frx":659A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Close this screen"
      Top             =   8520
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
End
Attribute VB_Name = "frmPurgeAir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub Update_PurgeDef()
    pnlPurge.Caption = "PurgeAir Source " & Format(DefPrg, "0")
    chkAkPurgeRequest.Value = IIf(Def_Prg.UsingPrgReqAK, 1, 0)
    chkHdwPurgeRequest.Value = IIf(Def_Prg.UsingPrgReqHdw, 1, 0)
    chkVacSw.Value = IIf(Def_Prg.UsingVacSwHdw, 1, 0)
    chkAuxAir.Value = IIf(Def_Prg.UsingAuxAirSol, 1, 0)
    chkPosPrsPrg.Value = IIf(Def_Prg.UsingPosPrsPrg, 1, 0)
End Sub

Sub Refresh_PurgeDef()
Dim func As Integer

    Def_Prg = PRG_INFO(DefPrg)
    txtDescription.text = Def_Prg.desc
    
    If USINGPRESSUREPURGE Then
        chkPosPrsPrg.Visible = True
    Else
        chkPosPrsPrg.Visible = False
        chkPosPrsPrg.Value = 0
    End If
    
    ' PurgeAir Analog Functions
    func = DefFunc(0, 3)
    txtaFuncDesc = Prg_AnaDef(func).desc
    txtaEUMax = Prg_AIO(DefPrg, func).EuMax
    txtaEUMin = Prg_AIO(DefPrg, func).EuMin
    txtaVDCMax = Prg_AIO(DefPrg, func).VdcMax
    txtaVDCMin = Prg_AIO(DefPrg, func).VdcMin
    txtaAddr = Prg_AIO(DefPrg, func).addr
    txtaChan = Prg_AIO(DefPrg, func).chan

    ' PurgeAir Digital Functions
    func = DefFunc(1, 3)
    txtdFuncDesc = Prg_DigDef(func).desc
    chkdInverse.Value = IIf(Prg_DIO(DefPrg, func).UseInverse, 1, 0)
    txtdAddr = Prg_DIO(DefPrg, func).addr
    txtdChan = Prg_DIO(DefPrg, func).chan

End Sub

Private Sub chkAkPurgeRequest_Click()
    Select Case chkAkPurgeRequest.Value
        Case 0
            ' unchecked
        Case 1
            ' checked
'            chkHdwPurgeRequest.Value = 0
    End Select
End Sub

Private Sub chkHdwPurgeRequest_Click()
    Select Case chkHdwPurgeRequest.Value
        Case 0
            ' unchecked
        Case 1
            ' checked
'            chkAkPurgeRequest.Value = 0
    End Select
End Sub

Private Sub cmdaDn_Click()
If DefFunc(0, 3) > 0 Then
    DefFunc(0, 3) = DefFunc(0, 3) - 1
Else
    DefFunc(0, 3) = MAX_ANA_PRG
End If
Refresh_PurgeDef
End Sub

Private Sub cmdaSave_Click()
Dim func As Integer

' PurgeAir Analog
func = DefFunc(0, 3)
Prg_AIO(DefPrg, func).EuMax = txtaEUMax
Prg_AIO(DefPrg, func).EuMin = txtaEUMin
Prg_AIO(DefPrg, func).VdcMax = txtaVDCMax
Prg_AIO(DefPrg, func).VdcMin = txtaVDCMin
Prg_AIO(DefPrg, func).addr = txtaAddr
Prg_AIO(DefPrg, func).chan = txtaChan
 
Save_AnalogFuncDef
Refresh_PurgeDef
End Sub

Private Sub cmdaUp_Click()
If DefFunc(0, 3) < MAX_ANA_PRG Then
    DefFunc(0, 3) = DefFunc(0, 3) + 1
Else
    DefFunc(0, 3) = 0
End If
Refresh_PurgeDef
End Sub

Private Sub cmddDn_Click()
If DefFunc(1, 3) > 0 Then
    DefFunc(1, 3) = DefFunc(1, 3) - 1
Else
    DefFunc(1, 3) = MAX_DIG_PRG
End If
Refresh_PurgeDef
End Sub

Private Sub cmdDown_Click()
    If DefPrg > 1 Then
        DefPrg = DefPrg - 1
    Else
        DefPrg = NR_PRGAIR
    End If
    Refresh_PurgeDef
    Update_PurgeDef
End Sub

Private Sub cmddSave_Click()
Dim func As Integer

' PurgeAir Digital
func = DefFunc(1, 3)
Prg_DIO(DefPrg, func).UseInverse = IIf(chkdInverse.Value, True, False)
Prg_DIO(DefPrg, func).addr = txtdAddr
Prg_DIO(DefPrg, func).chan = txtdChan
 
Save_DigitalFuncDef
Refresh_PurgeDef
End Sub

Private Sub cmddUp_Click()
If DefFunc(1, 3) < MAX_DIG_PRG Then
    DefFunc(1, 3) = DefFunc(1, 3) + 1
Else
    DefFunc(1, 3) = 0
End If
Refresh_PurgeDef
End Sub

Private Sub cmdExit_Click()

    Unload Me
    Set frmPurgeAir = Nothing
 
End Sub

Private Sub cmdSave_Click()

If CheckPass("9", True) Then

    Def_Prg.desc = txtDescription
    Def_Prg.CheckSecs = 3                                   ' hardcoded for now
    Def_Prg.UsingVacSwHdw = IIf(chkVacSw.Value = cYES, True, False)
    Def_Prg.UsingPrgReqAK = IIf(chkAkPurgeRequest.Value = cYES, True, False)
    Def_Prg.UsingPrgReqHdw = IIf(chkHdwPurgeRequest.Value = cYES, True, False)
    Def_Prg.UsingAuxAirSol = IIf(chkAuxAir.Value = cYES, True, False)
    Def_Prg.UsingPosPrsPrg = IIf(chkPosPrsPrg.Value = cYES, True, False)
    PRG_INFO(DefPrg) = Def_Prg
    
    ' Save PurgeAir Information
    Save_PurgeInfo
    
    Delay_Box "PurgeAir Source Info Saved", MSGDELAY, msgSHOW

End If
End Sub

Private Sub cmdSetDefaults_Click()
    Def_Prg.CheckSecs = 3
    Def_Prg.desc = "Common"
    Def_Prg.UsingAuxAirSol = True
    Def_Prg.UsingPosPrsPrg = False
    Def_Prg.UsingPrgReqAK = False
    Def_Prg.UsingPrgReqHdw = True
    Def_Prg.UsingVacSwHdw = False
    Update_PurgeDef
End Sub

Private Sub cmdUp_Click()
    If DefPrg < NR_PRGAIR Then
        DefPrg = DefPrg + 1
    Else
        DefPrg = 1
    End If
    Refresh_PurgeDef
    Update_PurgeDef
End Sub

Private Sub Form_Load()
    ' set foreground colors
    frmSelector.ForeColor = Titles_ForeColor
    pnlPurge.ForeColor = TitlesData_Forecolor
    frmPurge.ForeColor = Titles_ForeColor
    frmPurgeDef.ForeColor = TitlesLabel_ForeColor
    frmPrgFunc.ForeColor = TitlesLabel_ForeColor
    frmStnAnalog.ForeColor = TitlesData_Forecolor
    frmStnDigital.ForeColor = TitlesData_Forecolor
    ' refresh & update display
    Refresh_PurgeDef
    Update_PurgeDef
End Sub
