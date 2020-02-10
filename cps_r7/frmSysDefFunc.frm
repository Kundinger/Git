VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Begin VB.Form frmSysDefFunc 
   Caption         =   "Function Definition"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   Icon            =   "frmSysDefFunc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmCommon 
      Caption         =   "Common Functions"
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
      Height          =   4495
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   14655
      Begin VB.Frame frmCommonInputs 
         Caption         =   "Standard Analog Inputs"
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
         Height          =   1855
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   14425
         Begin VB.TextBox txtComAnaEUMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   95
            Text            =   "0"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   94
            Text            =   "0"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   93
            Text            =   "0"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   92
            Text            =   "1"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   91
            Text            =   "1"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   90
            Text            =   "1"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   89
            Text            =   "1"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaEUMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   88
            Text            =   "0"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   7905
            MaxLength       =   6
            TabIndex        =   87
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   6840
            MaxLength       =   6
            TabIndex        =   86
            Text            =   "1"
            ToolTipText     =   "Max Value in Volts"
            Top             =   1440
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   6840
            MaxLength       =   6
            TabIndex        =   85
            Text            =   "1"
            ToolTipText     =   "Max Value in Volts"
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   6840
            MaxLength       =   6
            TabIndex        =   84
            Text            =   "1"
            ToolTipText     =   "Max Value in Volts"
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMax 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   6840
            MaxLength       =   6
            TabIndex        =   83
            Text            =   "1"
            ToolTipText     =   "Max Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   7905
            MaxLength       =   6
            TabIndex        =   82
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   7905
            MaxLength       =   6
            TabIndex        =   81
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   840
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaVdcMin 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   7905
            MaxLength       =   6
            TabIndex        =   80
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtComAnaChan 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   79
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   1440
            Width           =   600
         End
         Begin VB.TextBox txtComAnaAddr 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   9180
            MaxLength       =   3
            TabIndex        =   78
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   1440
            Width           =   600
         End
         Begin VB.TextBox txtComAnaAddr 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   9180
            MaxLength       =   4
            TabIndex        =   77
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   1140
            Width           =   600
         End
         Begin VB.TextBox txtComAnaAddr 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   9180
            MaxLength       =   6
            TabIndex        =   76
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtComAnaAddr 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   9180
            MaxLength       =   3
            TabIndex        =   75
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtComAnaChan 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   74
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   1140
            Width           =   600
         End
         Begin VB.TextBox txtComAnaChan 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   9840
            MaxLength       =   5
            TabIndex        =   73
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   840
            Width           =   600
         End
         Begin VB.TextBox txtComAnaChan 
            Alignment       =   1  'Right Justify
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   72
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtComAnaDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   71
            Text            =   "1"
            ToolTipText     =   "Function Description"
            Top             =   1445
            Width           =   4125
         End
         Begin VB.TextBox txtComAnaDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   70
            Text            =   "1"
            ToolTipText     =   "Function Description"
            Top             =   1140
            Width           =   4125
         End
         Begin VB.TextBox txtComAnaDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   1
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   69
            Text            =   "1"
            ToolTipText     =   "Function Description"
            Top             =   840
            Width           =   4125
         End
         Begin VB.TextBox txtComAnaDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            DataField       =   "txtPATemp"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   68
            Text            =   "1"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4125
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
            Height          =   600
            Index           =   6
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":57E2
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Save Function Definition"
            Top             =   1115
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.CommandButton cmdComDefaults 
            Caption         =   "Set Common Defaults"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":5EE4
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Set Common Functions to Default Values"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Min"
            Height          =   255
            Left            =   5565
            TabIndex        =   106
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Max"
            Height          =   255
            Left            =   4560
            TabIndex        =   105
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Max"
            Height          =   255
            Left            =   6840
            TabIndex        =   104
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Min"
            Height          =   255
            Left            =   7905
            TabIndex        =   103
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Addr"
            Height          =   255
            Left            =   9180
            TabIndex        =   102
            Top             =   300
            Width           =   600
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Chan"
            Height          =   255
            Left            =   9840
            TabIndex        =   101
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   300
            Width           =   4005
         End
         Begin VB.Label lblUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            BackStyle       =   0  'Transparent
            Caption         =   "Temp in deg C"
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
            Height          =   285
            Index           =   0
            Left            =   10475
            TabIndex        =   99
            Top             =   565
            Width           =   1500
         End
         Begin VB.Label lblUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            BackStyle       =   0  'Transparent
            Caption         =   "Humidity in %"
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
            Height          =   285
            Index           =   1
            Left            =   10475
            TabIndex        =   98
            Top             =   865
            Width           =   1500
         End
         Begin VB.Label lblUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            BackStyle       =   0  'Transparent
            Caption         =   "Baro in mbar"
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
            Height          =   285
            Index           =   2
            Left            =   10475
            TabIndex        =   97
            Top             =   1165
            Width           =   1500
         End
         Begin VB.Label lblUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00000040&
            BackStyle       =   0  'Transparent
            Caption         =   "Leak in psig"
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
            Height          =   285
            Index           =   3
            Left            =   10475
            TabIndex        =   96
            Top             =   1465
            Width           =   1500
         End
      End
      Begin VB.Frame frmOtherAnalog 
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
         Index           =   0
         Left            =   120
         TabIndex        =   47
         Top             =   2200
         Width           =   14425
         Begin VB.TextBox txtaEUMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   57
            Text            =   "12345"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaEUMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   5565
            MaxLength       =   5
            TabIndex        =   56
            Text            =   "01234"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaVDCMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   6840
            MaxLength       =   8
            TabIndex        =   55
            Text            =   "12345678"
            ToolTipText     =   "Max Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaVDCMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   7905
            MaxLength       =   4
            TabIndex        =   54
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaAddr 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   9180
            MaxLength       =   2
            TabIndex        =   53
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtaChan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   52
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtaFuncDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   51
            Text            =   "Function Description123456789012345678901234567890"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4125
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
            Height          =   600
            Index           =   0
            Left            =   11160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":63D6
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
         End
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
            Height          =   600
            Index           =   0
            Left            =   10560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":6AD8
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "next"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
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
            Height          =   600
            Index           =   0
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":71DA
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Save Function Definition"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.Label lblaEUMax 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Max"
            Height          =   255
            Index           =   0
            Left            =   4560
            TabIndex        =   64
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblaEUMin 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Min"
            Height          =   255
            Index           =   0
            Left            =   5655
            TabIndex        =   63
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblaVdcMax 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Max"
            Height          =   255
            Index           =   0
            Left            =   6900
            TabIndex        =   62
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblaVdcMin 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Min"
            Height          =   255
            Index           =   0
            Left            =   7995
            TabIndex        =   61
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblaAddr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Addr"
            Height          =   255
            Index           =   0
            Left            =   9240
            TabIndex        =   60
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblaChan 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Chan"
            Height          =   255
            Index           =   0
            Left            =   9840
            TabIndex        =   59
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblaFuncDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   300
            Width           =   4005
         End
      End
      Begin VB.Frame frmOtherDigital 
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
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   3280
         Width           =   14415
         Begin VB.TextBox txtdChan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   43
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtdAddr 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   9180
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtdFuncDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   41
            Text            =   "Function Description123456789012345678901234567890"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4125
         End
         Begin VB.CheckBox chkdInverse 
            Alignment       =   1  'Right Justify
            Caption         =   "Use Inverse"
            Height          =   285
            Index           =   0
            Left            =   7500
            TabIndex        =   40
            Top             =   540
            Width           =   1335
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
            Height          =   600
            Index           =   0
            Left            =   11160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":78DC
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
         End
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
            Height          =   600
            Index           =   0
            Left            =   10560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":7FDE
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "next"
            Top             =   240
            UseMaskColor    =   -1  'True
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
            Height          =   600
            Index           =   0
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":86E0
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Save Function Definition"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.Label lbldFuncDesc 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   300
            Width           =   4005
         End
         Begin VB.Label lbldChan 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Chan"
            Height          =   255
            Index           =   0
            Left            =   9840
            TabIndex        =   45
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lbldAddr 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Addr"
            Height          =   255
            Index           =   0
            Left            =   9240
            TabIndex        =   44
            Top             =   300
            Width           =   555
         End
      End
   End
   Begin VB.Frame frmStation 
      Caption         =   "Station Functions"
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
      Height          =   2625
      Left            =   120
      TabIndex        =   5
      Top             =   4680
      Width           =   14655
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
         TabIndex        =   24
         Top             =   1440
         Width           =   14415
         Begin VB.TextBox txtdChan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   31
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtdAddr 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   9180
            MaxLength       =   2
            TabIndex        =   30
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtdFuncDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   29
            Text            =   "Function Description123456789012345678901234567890"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4125
         End
         Begin VB.CheckBox chkdInverse 
            Alignment       =   1  'Right Justify
            Caption         =   "Use Inverse"
            Height          =   285
            Index           =   2
            Left            =   7500
            TabIndex        =   28
            Top             =   540
            Width           =   1335
         End
         Begin VB.CommandButton cmddDn 
            DisabledPicture =   "frmSysDefFunc.frx":8DE2
            DownPicture     =   "frmSysDefFunc.frx":94E4
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   2
            Left            =   11160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":9BE6
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
         End
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
            Height          =   600
            Index           =   2
            Left            =   10560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":A2E8
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "next"
            Top             =   240
            UseMaskColor    =   -1  'True
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
            Height          =   600
            Index           =   2
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":A9EA
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Save Function Definition"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.Label lbldFuncDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   300
            Width           =   4005
         End
         Begin VB.Label lbldChan 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Chan"
            Height          =   255
            Index           =   2
            Left            =   9840
            TabIndex        =   33
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lbldAddr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Addr"
            Height          =   255
            Index           =   2
            Left            =   9240
            TabIndex        =   32
            Top             =   300
            Width           =   555
         End
      End
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
         TabIndex        =   6
         Top             =   300
         Width           =   14425
         Begin VB.TextBox txtaEUMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   4500
            MaxLength       =   6
            TabIndex        =   16
            Text            =   "12345"
            ToolTipText     =   "Max Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaEUMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   5565
            MaxLength       =   6
            TabIndex        =   15
            Text            =   "01234"
            ToolTipText     =   "Min Value in Engineering Units"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaVDCMax 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   6840
            MaxLength       =   6
            TabIndex        =   14
            Text            =   "12345678"
            ToolTipText     =   "Max Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaVDCMin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   7905
            MaxLength       =   6
            TabIndex        =   13
            Text            =   "0"
            ToolTipText     =   "Min Value in Volts"
            Top             =   540
            Width           =   1005
         End
         Begin VB.TextBox txtaAddr 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   9180
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "1"
            ToolTipText     =   "Opto Address (0-49)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtaChan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   9840
            MaxLength       =   4
            TabIndex        =   11
            Text            =   "0"
            ToolTipText     =   "Opto Channel (0-15)"
            Top             =   540
            Width           =   600
         End
         Begin VB.TextBox txtaFuncDesc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MaxLength       =   48
            TabIndex        =   10
            Text            =   "Function Description123456789012345678901234567890"
            ToolTipText     =   "Function Description"
            Top             =   540
            Width           =   4125
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
            Height          =   600
            Index           =   2
            Left            =   11160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":B0EC
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "previous"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
         End
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
            Height          =   600
            Index           =   2
            Left            =   10560
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":B7EE
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "next"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   600
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
            Height          =   600
            Index           =   2
            Left            =   12180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSysDefFunc.frx":BEF0
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Save Function Definition"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   2100
         End
         Begin VB.Label lblaEUMax 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Max"
            Height          =   255
            Index           =   2
            Left            =   4560
            TabIndex        =   23
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblaEUMin 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "EU Min"
            Height          =   255
            Index           =   2
            Left            =   5655
            TabIndex        =   22
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblaVdcMax 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Max"
            Height          =   255
            Index           =   2
            Left            =   6900
            TabIndex        =   21
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblaVdcMin 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Vdc Min"
            Height          =   255
            Index           =   2
            Left            =   7995
            TabIndex        =   20
            Top             =   300
            Width           =   885
         End
         Begin VB.Label lblaAddr 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Addr"
            Height          =   255
            Index           =   2
            Left            =   9240
            TabIndex        =   19
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblaChan 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Chan"
            Height          =   255
            Index           =   2
            Left            =   9840
            TabIndex        =   18
            Top             =   300
            Width           =   555
         End
         Begin VB.Label lblaFuncDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   4005
         End
      End
   End
   Begin VB.CommandButton cmdBack 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefFunc.frx":C5F2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Previous Screen"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdDown 
      DisabledPicture =   "frmSysDefFunc.frx":CCF4
      DownPicture     =   "frmSysDefFunc.frx":D3F6
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   10470
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefFunc.frx":DAF8
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdUp 
      DisabledPicture =   "frmSysDefFunc.frx":E73A
      DownPicture     =   "frmSysDefFunc.frx":EE3C
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   13935
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefFunc.frx":F53E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdStnDefaults 
      Caption         =   "Set Station Defaults"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   7800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefFunc.frx":10180
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Set Station Functions to Default Values"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   2280
   End
   Begin VB.CommandButton cmdStnClear 
      Caption         =   "Clear Station Definitions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSysDefFunc.frx":10672
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Clear Station Functions Definitions"
      Top             =   7440
      UseMaskColor    =   -1  'True
      Width           =   2280
   End
   Begin Threed.SSPanel pnlStn 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   645
      Left            =   11310
      TabIndex        =   107
      ToolTipText     =   "Station Number Displayed"
      Top             =   7440
      Width           =   2625
      _Version        =   65536
      _ExtentX        =   4630
      _ExtentY        =   1138
      _StockProps     =   15
      Caption         =   "49"
      ForeColor       =   -2147483646
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelInner      =   1
      Font3D          =   3
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "message"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   645
      Left            =   1680
      TabIndex        =   108
      Top             =   7440
      Width           =   3615
   End
End
Attribute VB_Name = "frmSysDefFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                           form frmSysDefFunc
Option Explicit
'
Private GoingToAnotherSysdef As Boolean

Sub Reset_Backgrounds()
Dim indx As Integer

    For indx = 0 To txtaAddr.UBound
        If (indx <> 1) Then txtaAddr(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtaChan.UBound
        If (indx <> 1) Then txtaChan(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtaEUMax.UBound
        If (indx <> 1) Then txtaEUMax(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtaEUMin.UBound
        If (indx <> 1) Then txtaEUMin(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtaVDCMax.UBound
        If (indx <> 1) Then txtaVDCMax(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtaVDCMin.UBound
        If (indx <> 1) Then txtaVDCMin(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtdAddr.UBound
        If (indx <> 1) Then txtdAddr(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To txtdChan.UBound
        If (indx <> 1) Then txtdChan(indx).BackColor = Entry_BackColor
    Next indx
    For indx = 0 To chkdInverse.UBound
        If (indx <> 1) Then chkdInverse(indx).ForeColor = Black
    Next indx
    lblMsg.Caption = " "
End Sub

Sub Refresh_FuncDef()
Dim func As Integer
Dim i As Integer

    pnlStn.Caption = "Station " & Format(DefStn, "0")
    
    ' Standard Common Analog
    i = 0
    func = acAmbTempSensor
    txtComAnaDesc(i) = Com_AnaDef(func).desc
    txtComAnaEUMax(i) = Com_AIO(func).EuMax
    txtComAnaEUMin(i) = Com_AIO(func).EuMin
    txtComAnaVdcMax(i) = Com_AIO(func).VdcMax
    txtComAnaVdcMin(i) = Com_AIO(func).VdcMin
    txtComAnaAddr(i) = Com_AIO(func).addr
    txtComAnaChan(i) = Com_AIO(func).chan
        
    i = 1
    func = acAmbHumiditySensor
    txtComAnaDesc(i) = Com_AnaDef(func).desc
    txtComAnaEUMax(i) = Com_AIO(func).EuMax
    txtComAnaEUMin(i) = Com_AIO(func).EuMin
    txtComAnaVdcMax(i) = Com_AIO(func).VdcMax
    txtComAnaVdcMin(i) = Com_AIO(func).VdcMin
    txtComAnaAddr(i) = Com_AIO(func).addr
    txtComAnaChan(i) = Com_AIO(func).chan
        
    i = 2
    func = acAmbBaroSensor
    txtComAnaDesc(i) = Com_AnaDef(func).desc
    txtComAnaEUMax(i) = Com_AIO(func).EuMax
    txtComAnaEUMin(i) = Com_AIO(func).EuMin
    txtComAnaVdcMax(i) = Com_AIO(func).VdcMax
    txtComAnaVdcMin(i) = Com_AIO(func).VdcMin
    txtComAnaAddr(i) = Com_AIO(func).addr
    txtComAnaChan(i) = Com_AIO(func).chan
        
    i = 3
    func = acComnPressSensor
    txtComAnaDesc(i) = Com_AnaDef(func).desc
    txtComAnaEUMax(i) = Com_AIO(func).EuMax
    txtComAnaEUMin(i) = Com_AIO(func).EuMin
    txtComAnaVdcMax(i) = Com_AIO(func).VdcMax
    txtComAnaVdcMin(i) = Com_AIO(func).VdcMin
    txtComAnaAddr(i) = Com_AIO(func).addr
    txtComAnaChan(i) = Com_AIO(func).chan
        
    ' Other Common Analog
    i = 0
    func = DefFunc(0, i)
    txtaFuncDesc(i) = Com_AnaDef(func).desc
    txtaEUMax(i) = Com_AIO(func).EuMax
    txtaEUMin(i) = Com_AIO(func).EuMin
    txtaVDCMax(i) = Com_AIO(func).VdcMax
    txtaVDCMin(i) = Com_AIO(func).VdcMin
    txtaAddr(i) = Com_AIO(func).addr
    txtaChan(i) = Com_AIO(func).chan
        
    If DefStn <> 0 Then
        ' Station Analog
        i = 2
        func = DefFunc(0, i)
        If (Stn_AnaDef(func).UsedIn(STN_INFO(DefStn).Type)) Then
            txtaFuncDesc(i) = Stn_AnaDef(func).desc
            txtaEUMax(i).Enabled = True
            txtaEUMin(i).Enabled = True
            txtaVDCMax(i).Enabled = True
            txtaVDCMin(i).Enabled = True
            txtaAddr(i).Enabled = True
            txtaChan(i).Enabled = True
            txtaEUMax(i) = Stn_AIO(DefStn, func).EuMax
            txtaEUMin(i) = Stn_AIO(DefStn, func).EuMin
            txtaVDCMax(i) = Stn_AIO(DefStn, func).VdcMax
            txtaVDCMin(i) = Stn_AIO(DefStn, func).VdcMin
            txtaAddr(i) = Stn_AIO(DefStn, func).addr
            txtaChan(i) = Stn_AIO(DefStn, func).chan
        Else
            txtaFuncDesc(i) = "stn analog unused " & Format(func, "00")
            txtaEUMax(i).Enabled = False
            txtaEUMin(i).Enabled = False
            txtaVDCMax(i).Enabled = False
            txtaVDCMin(i).Enabled = False
            txtaAddr(i).Enabled = False
            txtaChan(i).Enabled = False
            txtaEUMax(i) = Stn_AIO(DefStn, func).EuMax
            txtaEUMin(i) = Stn_AIO(DefStn, func).EuMin
            txtaVDCMax(i) = Stn_AIO(DefStn, func).VdcMax
            txtaVDCMin(i) = Stn_AIO(DefStn, func).VdcMin
            txtaAddr(i) = Stn_AIO(DefStn, func).addr
            txtaChan(i) = Stn_AIO(DefStn, func).chan
        End If
    End If
    
    ' Common Digital
    i = 0
    func = DefFunc(1, i)
    txtdFuncDesc(i) = Com_DigDef(func).desc
    chkdInverse(i).Value = IIf(Com_DIO(func).UseInverse, 1, 0)
    txtdAddr(i) = Com_DIO(func).addr
    txtdChan(i) = Com_DIO(func).chan
    
    ' Station Digital
    i = 2
    func = DefFunc(1, i)
    If (Stn_DigDef(func).UsedIn(STN_INFO(DefStn).Type)) Then
        txtdFuncDesc(i) = Stn_DigDef(func).desc
        chkdInverse(i).Enabled = True
        txtdAddr(i).Enabled = True
        txtdChan(i).Enabled = True
        chkdInverse(i).Value = IIf(Stn_DIO(DefStn, func).UseInverse, 1, 0)
        txtdAddr(i) = Stn_DIO(DefStn, func).addr
        txtdChan(i) = Stn_DIO(DefStn, func).chan
    Else
        txtdFuncDesc(i) = "stn digital unused " & Format(func, "00")
        chkdInverse(i).Enabled = False
        txtdAddr(i).Enabled = False
        txtdChan(i).Enabled = False
        chkdInverse(i).Value = IIf(Stn_DIO(DefStn, func).UseInverse, 1, 0)
        txtdAddr(i) = Stn_DIO(DefStn, func).addr
        txtdChan(i) = Stn_DIO(DefStn, func).chan
    End If
    
    Reset_Backgrounds
End Sub

Private Sub chkdInverse_Click(Index As Integer)
    chkdInverse(Index).ForeColor = DKPURPLE
End Sub

Private Sub cmdaDn_Click(Index As Integer)
Dim max As Integer
    Select Case Index
        Case 0
            max = MAX_ANA_COM
        Case 2
            max = MAX_ANA_STN
        Case Else
            ' do nothing
    End Select
    If DefFunc(0, Index) > 0 Then
        DefFunc(0, Index) = DefFunc(0, Index) - 1
    Else
        DefFunc(0, Index) = max
    End If
    Refresh_FuncDef
End Sub

Private Sub cmdaSave_Click(Index As Integer)
Dim func As Integer
Dim i As Integer

    Select Case Index
        Case 0
            ' Other Common Analog
            func = DefFunc(0, Index)
            Com_AIO(func).EuMax = txtaEUMax(Index)
            Com_AIO(func).EuMin = txtaEUMin(Index)
            Com_AIO(func).VdcMax = txtaVDCMax(Index)
            Com_AIO(func).VdcMin = txtaVDCMin(Index)
            Com_AIO(func).addr = txtaAddr(Index)
            Com_AIO(func).chan = txtaChan(Index)
     
        Case 2
            ' Station Analog
            func = DefFunc(0, Index)
            Stn_AIO(DefStn, func).EuMax = txtaEUMax(Index)
            Stn_AIO(DefStn, func).EuMin = txtaEUMin(Index)
            Stn_AIO(DefStn, func).VdcMax = txtaVDCMax(Index)
            Stn_AIO(DefStn, func).VdcMin = txtaVDCMin(Index)
            Stn_AIO(DefStn, func).addr = txtaAddr(Index)
            Stn_AIO(DefStn, func).chan = txtaChan(Index)
     
        Case 3
            ' Standard Common Analog
            i = 0
            func = acAmbTempSensor
            Com_AIO(func).EuMax = txtComAnaEUMax(i)
            Com_AIO(func).EuMin = txtComAnaEUMin(i)
            Com_AIO(func).VdcMax = txtComAnaVdcMax(i)
            Com_AIO(func).VdcMin = txtComAnaVdcMin(i)
            Com_AIO(func).addr = txtComAnaAddr(i)
            Com_AIO(func).chan = txtComAnaChan(i)
            
            i = 1
            func = acAmbHumiditySensor
            Com_AIO(func).EuMax = txtComAnaEUMax(i)
            Com_AIO(func).EuMin = txtComAnaEUMin(i)
            Com_AIO(func).VdcMax = txtComAnaVdcMax(i)
            Com_AIO(func).VdcMin = txtComAnaVdcMin(i)
            Com_AIO(func).addr = txtComAnaAddr(i)
            Com_AIO(func).chan = txtComAnaChan(i)
                
            i = 2
            func = acAmbBaroSensor
            Com_AIO(func).EuMax = txtComAnaEUMax(i)
            Com_AIO(func).EuMin = txtComAnaEUMin(i)
            Com_AIO(func).VdcMax = txtComAnaVdcMax(i)
            Com_AIO(func).VdcMin = txtComAnaVdcMin(i)
            Com_AIO(func).addr = txtComAnaAddr(i)
            Com_AIO(func).chan = txtComAnaChan(i)
                
            i = 3
            func = acComnPressSensor
            Com_AIO(func).EuMax = txtComAnaEUMax(i)
            Com_AIO(func).EuMin = txtComAnaEUMin(i)
            Com_AIO(func).VdcMax = txtComAnaVdcMax(i)
            Com_AIO(func).VdcMin = txtComAnaVdcMin(i)
            Com_AIO(func).addr = txtComAnaAddr(i)
            Com_AIO(func).chan = txtComAnaChan(i)
                
     
    End Select
    Save_AnalogFuncDef
    Refresh_FuncDef
End Sub

Private Sub cmdaUp_Click(Index As Integer)
Dim max As Integer
    Select Case Index
        Case 0
            max = MAX_ANA_COM
        Case 2
            max = MAX_ANA_STN
        Case Else
            ' do nothing
    End Select
    If DefFunc(0, Index) < max Then
        DefFunc(0, Index) = DefFunc(0, Index) + 1
    Else
        DefFunc(0, Index) = 0
    End If
    Refresh_FuncDef
End Sub

Private Sub cmdBack_Click()
    frmSysDefStn.Show
    GoingToAnotherSysdef = True
    Unload Me
End Sub

Private Sub cmdComDefaults_Click()

Dim chan As Integer
Dim func As Integer
Dim baseaddr As Integer

    If CheckPass("9", True) Then
    
        ' Base Opto Address for Commmon Board
        baseaddr = 0
        
        
        ' Common Digital
        func = icHornSilencePB
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 0
         
        func = icExhaustFlowFS
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 1
         
        func = icEStopSw
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 2
         
        func = ic20LelGasSw
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 6
         
        func = icAlarmBeacon
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 13
         
        func = icAlarmHorn
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 14
         
        func = icPauseLT
        Com_DIO(func).UseInverse = False
        Com_DIO(func).addr = baseaddr
        Com_DIO(func).chan = 15
         
        
        ' Save the Digital Changes
        Save_DigitalFuncDef
        
        
        ' Common Analogs
        
        func = acAmbTempSensor
        Com_AIO(func).EuMax = 140
        Com_AIO(func).EuMin = -20
        Com_AIO(func).VdcMax = 5
        Com_AIO(func).VdcMin = 1
        Com_AIO(func).addr = baseaddr + 3
        Com_AIO(func).chan = 4
              
        func = acAmbHumiditySensor
        Com_AIO(func).EuMax = 100
        Com_AIO(func).EuMin = 0
        Com_AIO(func).VdcMax = 5
        Com_AIO(func).VdcMin = 1
        Com_AIO(func).addr = baseaddr + 3
        Com_AIO(func).chan = 5
              
        func = acComnPressSensor
        Com_AIO(func).EuMax = 15
        Com_AIO(func).EuMin = 0
        Com_AIO(func).VdcMax = 5
        Com_AIO(func).VdcMin = 1
        Com_AIO(func).addr = baseaddr + 3
        Com_AIO(func).chan = 6
              
        func = acAmbBaroSensor
        Com_AIO(func).EuMax = 15.9
        Com_AIO(func).EuMin = 8.7
        Com_AIO(func).VdcMax = 5
        Com_AIO(func).VdcMin = 0
        Com_AIO(func).addr = baseaddr + 3
        Com_AIO(func).chan = 7
              
              
        ' Save the AnalogChanges
        Save_AnalogFuncDef
        
        Refresh_FuncDef
        
        Delay_Box "Station Function Configuration set to Defaults", MSGDELAY, msgSHOW
        lblMsg.Caption = "Common Function Configuration set to Defaults"
            
    End If
End Sub


Private Sub cmddDn_Click(Index As Integer)
Dim max As Integer
    Select Case Index
        Case 0
            max = MAX_DIG_COM
        Case 2
            max = MAX_DIG_STN
        Case Else
            ' do nothing
    End Select
    If DefFunc(1, Index) > 0 Then
        DefFunc(1, Index) = DefFunc(1, Index) - 1
    Else
        DefFunc(1, Index) = max
    End If
    Refresh_FuncDef
End Sub

Private Sub cmdDown_Click()
    If DefStn > 1 Then
        DefStn = DefStn - 1
    Else
        DefStn = NR_STN
    End If
    Refresh_FuncDef
End Sub

Private Sub cmddSave_Click(Index As Integer)
Dim func As Integer

    Select Case Index
        Case 0
            ' Common Digital
            func = DefFunc(1, Index)
            Com_DIO(func).UseInverse = IIf(chkdInverse(Index).Value, True, False)
            Com_DIO(func).addr = txtdAddr(Index)
            Com_DIO(func).chan = txtdChan(Index)
     
        Case 2
            ' Station Digital
            func = DefFunc(1, Index)
            Stn_DIO(DefStn, func).UseInverse = IIf(chkdInverse(Index).Value, True, False)
            Stn_DIO(DefStn, func).addr = txtdAddr(Index)
            Stn_DIO(DefStn, func).chan = txtdChan(Index)
     
    End Select
    Save_DigitalFuncDef
    Refresh_FuncDef
End Sub

Private Sub cmddUp_Click(Index As Integer)
Dim max As Integer
    Select Case Index
        Case 0
            max = MAX_DIG_COM
        Case 2
            max = MAX_DIG_STN
        Case Else
            ' do nothing
    End Select
    If DefFunc(1, Index) < max Then
        DefFunc(1, Index) = DefFunc(1, Index) + 1
    Else
        DefFunc(1, Index) = 0
    End If
    Refresh_FuncDef
End Sub

Private Sub cmdStnClear_Click()
Dim chan As Integer
Dim func As Integer
Dim baseaddr As Integer

    If CheckPass("9", True) Then
    
        ' Base Opto Address for Current Station
        baseaddr = DefStn * 4
        
        
        ' Station Digitals
        For func = 1 To MAX_DIG_STN
            Stn_DIO(DefStn, func).UseInverse = False
            Stn_DIO(DefStn, func).addr = 0
            Stn_DIO(DefStn, func).chan = 0
        Next func
         
        ' Save the Digital Changes
        Save_DigitalFuncDef
        
        
        ' Station Analogs
        For func = 1 To MAX_ANA_STN
            Stn_AIO(DefStn, func).EuMax = 0
            Stn_AIO(DefStn, func).EuMin = 0
            Stn_AIO(DefStn, func).VdcMax = 0
            Stn_AIO(DefStn, func).VdcMin = 0
            Stn_AIO(DefStn, func).addr = 0
            Stn_AIO(DefStn, func).chan = 0
        Next func
         
        ' Save the Analog Changes
        Save_AnalogFuncDef
        
        Refresh_FuncDef
        
        Delay_Box "Station Function Configuration Cleared", MSGDELAY, msgSHOW
        lblMsg.Caption = "Station Function Configuration Cleared"
        
    End If
End Sub

Private Sub cmdStnDefaults_Click()
Dim chan As Integer
Dim func As Integer
Dim baseaddr As Integer

    If CheckPass("9", True) Then
    
        ' Base Opto Address for Current Station
        baseaddr = DefStn * 4
        
        
        ' Station Digitals
        
        func = isNitrogenSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 0
         
        func = isButaneSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 1
         
        func = isPurgeSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 2
         
        func = isPriDirectionSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 3
         
        func = isAuxCanVentSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 4
         
        func = isLeakCheckSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 5
         
        func = isAuxPurgeSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 6
         
        func = isPriAuxVentSol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 7
         
        func = isLoadShift2Sol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 8
         
        func = isVentShift2Sol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 9
         
        func = isPurgeShift2Sol
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 10
         
        func = isPauseLT
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 11
         
        func = isIdleLT
        Stn_DIO(DefStn, func).UseInverse = False
        Stn_DIO(DefStn, func).addr = baseaddr
        Stn_DIO(DefStn, func).chan = 12
     
        ' Save the Digital Changes
        Save_DigitalFuncDef
        
        ' Station Analogs
        
        func = asNitrogenFlowSP
        Stn_AIO(DefStn, func).EuMax = 1
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 2
        Stn_AIO(DefStn, func).chan = 12
         
        func = asButaneFlowSP
        Stn_AIO(DefStn, func).EuMax = 1
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 2
        Stn_AIO(DefStn, func).chan = 13
         
        func = asPurgeAirFlowSP
        Stn_AIO(DefStn, func).EuMax = 30
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 2
        Stn_AIO(DefStn, func).chan = 14
         
        func = asNitrogenFlow
        Stn_AIO(DefStn, func).EuMax = 1
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 3
        Stn_AIO(DefStn, func).chan = 2
         
        func = asButaneFlow
        Stn_AIO(DefStn, func).EuMax = 1
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 3
        Stn_AIO(DefStn, func).chan = 3
         
        func = asPurgeAirFlow
        Stn_AIO(DefStn, func).EuMax = 30
        Stn_AIO(DefStn, func).EuMin = 0
        Stn_AIO(DefStn, func).VdcMax = 5
        Stn_AIO(DefStn, func).VdcMin = 0
        Stn_AIO(DefStn, func).addr = baseaddr + 3
        Stn_AIO(DefStn, func).chan = 4
         
        ' Save the Analog Changes
        Save_AnalogFuncDef
        
        Refresh_FuncDef
        
        Delay_Box "Station Function Configuration set to Defaults", MSGDELAY, msgSHOW
        lblMsg.Caption = "Station Function Configuration set to Defaults"
        
    End If
End Sub

Private Sub cmdUp_Click()
    If DefStn < NR_STN Then
        DefStn = DefStn + 1
    Else
        DefStn = 1
    End If
    Refresh_FuncDef
End Sub

Private Sub Form_Load()
Dim indx As Integer
    GoingToAnotherSysdef = False
'    If Not ReadyToRun Then
'        lblMsg.ForeColor = lblMsg.BackColor
'    Else
        lblMsg.ForeColor = Message_ForeColor
'    End If
    ' Set Title Foreground color
    frmCommon.ForeColor = Titles_ForeColor
    frmCommonInputs.ForeColor = TitlesData_Forecolor
    frmOtherAnalog(0).ForeColor = TitlesData_Forecolor
    frmOtherDigital(0).ForeColor = TitlesData_Forecolor
    frmStation.ForeColor = Titles_ForeColor
    frmStnAnalog.ForeColor = TitlesData_Forecolor
    frmStnDigital.ForeColor = TitlesData_Forecolor
    pnlStn.ForeColor = TitlesData_Forecolor
    ' set background colors
    pnlStn.BackColor = EntryNotChangeable_BackColor
    For indx = 0 To txtComAnaDesc.UBound
        txtComAnaDesc(indx).BackColor = Common_BackColor
    Next indx
    For indx = 0 To txtaFuncDesc.UBound
        If (indx <> 1) Then txtaFuncDesc(indx).BackColor = Common_BackColor
    Next indx
    For indx = 0 To txtdFuncDesc.UBound
        If (indx <> 1) Then txtdFuncDesc(indx).BackColor = Common_BackColor
    Next indx

    
    If DefStn = 0 Then DefStn = 1
    If USINGC Then
        lblUnits(0).Caption = "Temp in deg C"
    Else
        lblUnits(0).Caption = "Temp in deg F"
    End If
    lblUnits(1).Caption = "Humidity in %"
    lblUnits(2).Caption = "Baro in mbar"
    lblUnits(3).Caption = "Leak in psig"
    
    lblMsg.Caption = " "
    
    Form_Center Me
    
    Refresh_FuncDef
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not GoingToAnotherSysdef Then ReadyToRun = True
    Unload Me
End Sub

Private Sub txtaAddr_Change(Index As Integer)
    txtaAddr(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtaChan_Change(Index As Integer)
    txtaChan(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtaEUMax_Change(Index As Integer)
    txtaEUMax(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtaEUMin_Change(Index As Integer)
    txtaEUMin(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtaVDCMax_Change(Index As Integer)
    txtaVDCMax(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtaVDCMin_Change(Index As Integer)
    txtaVDCMin(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtdAddr_Change(Index As Integer)
    txtdAddr(Index).BackColor = PALEYELLOW
End Sub

Private Sub txtdChan_Change(Index As Integer)
    txtdChan(Index).BackColor = PALEYELLOW
End Sub
