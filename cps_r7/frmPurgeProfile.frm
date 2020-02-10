VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Begin VB.Form frmPurgeProfile 
   BackColor       =   &H00C000C0&
   Caption         =   "Set Point Versus Time Profile"
   ClientHeight    =   10620
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12210
   ControlBox      =   0   'False
   Icon            =   "frmPurgeProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotHighlight 
      Height          =   285
      Left            =   1680
      TabIndex        =   181
      Text            =   "NOT Highlight"
      Top             =   11880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmPage 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   10635
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12200
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         DisabledPicture =   "frmPurgeProfile.frx":57E2
         DownPicture     =   "frmPurgeProfile.frx":5EE4
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":65E6
         Style           =   1  'Graphical
         TabIndex        =   185
         ToolTipText     =   "Import Profile from Import.prg file in Recipes folder"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         DisabledPicture =   "frmPurgeProfile.frx":7228
         DownPicture     =   "frmPurgeProfile.frx":792A
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":802C
         Style           =   1  'Graphical
         TabIndex        =   183
         ToolTipText     =   "Copy Profile Steps to the clipboard"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         DisabledPicture =   "frmPurgeProfile.frx":872E
         DownPicture     =   "frmPurgeProfile.frx":8E30
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":9532
         Style           =   1  'Graphical
         TabIndex        =   182
         ToolTipText     =   "Paste Profile Steps from the clipboard"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Frame frmStatus 
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
         Height          =   1755
         Left            =   120
         TabIndex        =   11
         Top             =   8760
         Width           =   11955
         Begin VB.Label lblMessage 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "messages"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   11730
         End
      End
      Begin VB.Frame frmSteps 
         Caption         =   "Steps"
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
         Height          =   5895
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   11955
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   14
            Left            =   10875
            TabIndex        =   179
            Top             =   2174
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   13
            Left            =   10875
            TabIndex        =   178
            Top             =   1792
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   12
            Left            =   10875
            TabIndex        =   177
            Top             =   1410
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   11
            Left            =   10875
            TabIndex        =   176
            Top             =   1028
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   15
            Left            =   10875
            TabIndex        =   175
            Top             =   2556
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   16
            Left            =   10875
            TabIndex        =   174
            Top             =   2938
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   17
            Left            =   10875
            TabIndex        =   173
            Top             =   3320
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   18
            Left            =   10875
            TabIndex        =   172
            Top             =   3702
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   19
            Left            =   10875
            TabIndex        =   171
            Top             =   4084
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   20
            Left            =   10875
            TabIndex        =   170
            Top             =   4470
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   4
            Left            =   4845
            TabIndex        =   166
            Top             =   2178
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   3
            Left            =   4845
            TabIndex        =   165
            Top             =   1797
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   2
            Left            =   4845
            TabIndex        =   164
            Top             =   1416
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   1
            Left            =   4845
            TabIndex        =   163
            Top             =   1035
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   5
            Left            =   4845
            TabIndex        =   162
            Top             =   2559
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   6
            Left            =   4845
            TabIndex        =   161
            Top             =   2940
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   7
            Left            =   4845
            TabIndex        =   160
            Top             =   3321
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   8
            Left            =   4845
            TabIndex        =   159
            Top             =   3702
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   9
            Left            =   4845
            TabIndex        =   158
            Top             =   4083
            Width           =   255
         End
         Begin VB.CheckBox chkLastStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   10
            Left            =   4845
            TabIndex        =   157
            Top             =   4470
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   20
            Left            =   8370
            TabIndex        =   153
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   4425
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   20
            Left            =   7350
            TabIndex        =   152
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   4425
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   20
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   151
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   4440
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   20
            Left            =   9630
            TabIndex        =   150
            Top             =   4470
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   20
            Left            =   10230
            TabIndex        =   149
            Top             =   4470
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   19
            Left            =   8370
            TabIndex        =   148
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   4035
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   19
            Left            =   7350
            TabIndex        =   147
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   4035
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   19
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   146
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   4050
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   19
            Left            =   9630
            TabIndex        =   145
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   19
            Left            =   10230
            TabIndex        =   144
            Top             =   4080
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   18
            Left            =   8370
            TabIndex        =   143
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   3660
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   18
            Left            =   7350
            TabIndex        =   142
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   3660
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   18
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   141
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   3675
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   18
            Left            =   9630
            TabIndex        =   140
            Top             =   3705
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   18
            Left            =   10230
            TabIndex        =   139
            Top             =   3705
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   17
            Left            =   8370
            TabIndex        =   138
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   3285
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   17
            Left            =   7350
            TabIndex        =   137
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   3285
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   17
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   136
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   3300
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   17
            Left            =   9630
            TabIndex        =   135
            Top             =   3330
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   17
            Left            =   10230
            TabIndex        =   134
            Top             =   3330
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   16
            Left            =   8370
            TabIndex        =   133
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2895
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   16
            Left            =   7350
            TabIndex        =   132
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2895
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   16
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   131
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2910
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   16
            Left            =   9630
            TabIndex        =   130
            Top             =   2940
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   16
            Left            =   10230
            TabIndex        =   129
            Top             =   2940
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   15
            Left            =   8370
            TabIndex        =   128
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2520
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   15
            Left            =   7350
            TabIndex        =   127
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2520
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   15
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   126
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2535
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   15
            Left            =   9630
            TabIndex        =   125
            Top             =   2565
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   15
            Left            =   10200
            TabIndex        =   124
            Top             =   2565
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   14
            Left            =   8370
            TabIndex        =   123
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2145
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   14
            Left            =   7350
            TabIndex        =   122
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2145
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   14
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   121
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2160
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   14
            Left            =   9630
            TabIndex        =   120
            Top             =   2190
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   14
            Left            =   10230
            TabIndex        =   119
            Top             =   2190
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   13
            Left            =   8370
            TabIndex        =   118
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   1755
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   13
            Left            =   7350
            TabIndex        =   117
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   1755
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   13
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   116
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1770
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   13
            Left            =   9630
            TabIndex        =   115
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   13
            Left            =   10230
            TabIndex        =   114
            Top             =   1800
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   12
            Left            =   8370
            TabIndex        =   113
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   1380
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   12
            Left            =   7350
            TabIndex        =   112
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   1380
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   12
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   111
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1395
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   12
            Left            =   9630
            TabIndex        =   110
            Top             =   1425
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   12
            Left            =   10230
            TabIndex        =   109
            Top             =   1425
            Width           =   255
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   11
            Left            =   8370
            TabIndex        =   108
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   990
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   11
            Left            =   7350
            TabIndex        =   107
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   990
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   11
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   106
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1005
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   11
            Left            =   9630
            TabIndex        =   105
            Top             =   1028
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   11
            Left            =   10230
            TabIndex        =   104
            Top             =   1028
            Width           =   255
         End
         Begin VB.CommandButton cmdStepPageDn 
            Caption         =   " Prev 10 Steps"
            DisabledPicture =   "frmPurgeProfile.frx":9C34
            DownPicture     =   "frmPurgeProfile.frx":A336
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   4440
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":AA38
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   1440
         End
         Begin VB.CommandButton cmdStepPageUp 
            Caption         =   " Next 10 Steps"
            DisabledPicture =   "frmPurgeProfile.frx":B67A
            DownPicture     =   "frmPurgeProfile.frx":BD7C
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   6120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":C47E
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   1440
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   1
            Left            =   8520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":D0C0
            Style           =   1  'Graphical
            TabIndex        =   101
            ToolTipText     =   "Clear Profile Steps"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton cmdPaste10Steps 
            Caption         =   "Paste"
            DisabledPicture =   "frmPurgeProfile.frx":DD02
            DownPicture     =   "frmPurgeProfile.frx":E404
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   1
            Left            =   10470
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":EB06
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "Paste Profile Steps from the clipboard"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton cmdCopy10Steps 
            Caption         =   "Copy"
            DisabledPicture =   "frmPurgeProfile.frx":F208
            DownPicture     =   "frmPurgeProfile.frx":F90A
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   1
            Left            =   9630
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":1000C
            Style           =   1  'Graphical
            TabIndex        =   99
            ToolTipText     =   "Copy Profile Steps to the clipboard"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   10
            Left            =   3600
            TabIndex        =   90
            Top             =   4470
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   10
            Left            =   4200
            TabIndex        =   89
            Top             =   4470
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   9
            Left            =   3600
            TabIndex        =   88
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   9
            Left            =   4200
            TabIndex        =   87
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   8
            Left            =   3600
            TabIndex        =   86
            Top             =   3705
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   8
            Left            =   4200
            TabIndex        =   85
            Top             =   3705
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   7
            Left            =   3600
            TabIndex        =   84
            Top             =   3330
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   7
            Left            =   4200
            TabIndex        =   83
            Top             =   3330
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   6
            Left            =   3600
            TabIndex        =   82
            Top             =   2940
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   6
            Left            =   4200
            TabIndex        =   81
            Top             =   2940
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   5
            Left            =   3600
            TabIndex        =   80
            Top             =   2565
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   79
            Top             =   2565
            Width           =   255
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   0
            Left            =   690
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":1070E
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Clear Profile Steps"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton cmdPaste10Steps 
            Caption         =   "Paste"
            DisabledPicture =   "frmPurgeProfile.frx":11350
            DownPicture     =   "frmPurgeProfile.frx":11A52
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   0
            Left            =   2640
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":12154
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Paste Profile Steps from the clipboard"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.CommandButton cmdCopy10Steps 
            Caption         =   "Copy"
            DisabledPicture =   "frmPurgeProfile.frx":12856
            DownPicture     =   "frmPurgeProfile.frx":12F58
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Index           =   0
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":1365A
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Copy Profile Steps to the clipboard"
            Top             =   4860
            UseMaskColor    =   -1  'True
            Width           =   840
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   10
            Left            =   2340
            TabIndex        =   63
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   4425
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   9
            Left            =   2340
            TabIndex        =   62
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   4035
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   8
            Left            =   2340
            TabIndex        =   61
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   3660
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   7
            Left            =   2340
            TabIndex        =   60
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   3285
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   6
            Left            =   2340
            TabIndex        =   59
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2895
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   5
            Left            =   2340
            TabIndex        =   58
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2520
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   4
            Left            =   2340
            TabIndex        =   57
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   2145
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   3
            Left            =   2340
            TabIndex        =   56
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   1755
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   2
            Left            =   2340
            TabIndex        =   55
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   1380
            Width           =   900
         End
         Begin VB.TextBox txtStepDuration 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   1
            Left            =   2340
            TabIndex        =   54
            Text            =   "Time"
            ToolTipText     =   "Time in minutes"
            Top             =   990
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   1
            Left            =   1320
            TabIndex        =   53
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   1005
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   2
            Left            =   1320
            TabIndex        =   52
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   1380
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   3
            Left            =   1320
            TabIndex        =   51
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   1755
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   4
            Left            =   1320
            TabIndex        =   50
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2145
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   5
            Left            =   1320
            TabIndex        =   49
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2520
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   6
            Left            =   1320
            TabIndex        =   48
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   2895
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   7
            Left            =   1320
            TabIndex        =   47
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   3285
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   8
            Left            =   1320
            TabIndex        =   46
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   3660
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   9
            Left            =   1320
            TabIndex        =   45
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   4035
            Width           =   900
         End
         Begin VB.TextBox txtInitialSP 
            Alignment       =   1  'Right Justify
            Height          =   330
            Index           =   10
            Left            =   1320
            TabIndex        =   44
            Text            =   "st p"
            ToolTipText     =   "Enter Flow in slpm"
            Top             =   4425
            Width           =   900
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   1
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   43
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1005
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   2
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   42
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1380
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   3
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   1755
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   4
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2145
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   5
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   39
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2520
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   6
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   38
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   2895
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   7
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   3285
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   8
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   36
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   3660
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   9
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   35
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   4035
            Width           =   555
         End
         Begin VB.TextBox txtStepNumber 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            Height          =   300
            Index           =   10
            Left            =   690
            Locked          =   -1  'True
            TabIndex        =   34
            Text            =   "Text1"
            ToolTipText     =   "Purge Profile Step Number"
            Top             =   4425
            Width           =   555
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   33
            Top             =   1035
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   32
            Top             =   1035
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   2
            Left            =   3600
            TabIndex        =   31
            Top             =   1425
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   30
            Top             =   1425
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   29
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   28
            Top             =   1800
            Width           =   255
         End
         Begin VB.CheckBox chkStepStep 
            Caption         =   "step"
            Height          =   255
            Index           =   4
            Left            =   3600
            TabIndex        =   27
            Top             =   2190
            Width           =   255
         End
         Begin VB.CheckBox chkRampStep 
            Caption         =   "ramp"
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   26
            Top             =   2190
            Width           =   255
         End
         Begin VB.Label lblLastStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   1
            Left            =   10680
            TabIndex        =   180
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblRampStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "ramp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   1
            Left            =   10035
            TabIndex        =   169
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblStepStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   1
            Left            =   9390
            TabIndex        =   168
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblLastStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "last"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   0
            Left            =   4650
            TabIndex        =   167
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblStepType 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   9480
            TabIndex        =   156
            Top             =   465
            Width           =   1620
         End
         Begin VB.Label lblStepTypeDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   9480
            TabIndex        =   155
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label lblStepType 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   3510
            TabIndex        =   154
            Top             =   525
            Width           =   1740
         End
         Begin VB.Label lblDuration 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   8370
            TabIndex        =   98
            Top             =   465
            Width           =   900
         End
         Begin VB.Label lblSP 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Set Pt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   7350
            TabIndex        =   97
            Top             =   465
            Width           =   900
         End
         Begin VB.Label lblSpDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Initial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   7350
            TabIndex        =   96
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   1
            Left            =   6720
            TabIndex        =   95
            Top             =   465
            Width           =   555
         End
         Begin VB.Label lblSpUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(slpm)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   7350
            TabIndex        =   94
            Top             =   690
            Width           =   900
         End
         Begin VB.Label lblDurationUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(min)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   8370
            TabIndex        =   93
            Top             =   690
            Width           =   900
         End
         Begin VB.Label lblStepUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   6720
            TabIndex        =   92
            Top             =   690
            Width           =   555
         End
         Begin VB.Label lblDurationDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   1
            Left            =   8370
            TabIndex        =   91
            Top             =   240
            Width           =   900
         End
         Begin VB.Label lblDuration 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   2340
            TabIndex        =   74
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lblSP 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Set Pt"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   1320
            TabIndex        =   73
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lblSpDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Initial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   1320
            TabIndex        =   72
            Top             =   300
            Width           =   900
         End
         Begin VB.Label lblStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   0
            Left            =   690
            TabIndex        =   71
            Top             =   525
            Width           =   555
         End
         Begin VB.Label lblSpUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(slpm)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   1320
            TabIndex        =   70
            Top             =   750
            Width           =   900
         End
         Begin VB.Label lblDurationUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "(min)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   2340
            TabIndex        =   69
            Top             =   750
            Width           =   900
         End
         Begin VB.Label lblStepUnits 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "#"
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
            Height          =   270
            Index           =   0
            Left            =   690
            TabIndex        =   68
            Top             =   750
            Width           =   555
         End
         Begin VB.Label lblDurationDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   2340
            TabIndex        =   67
            Top             =   300
            Width           =   900
         End
         Begin VB.Label lblStepStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   0
            Left            =   3360
            TabIndex        =   66
            Top             =   735
            Width           =   645
         End
         Begin VB.Label lblStepTypeDesc 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Index           =   0
            Left            =   3510
            TabIndex        =   65
            Top             =   300
            Width           =   1740
         End
         Begin VB.Label lblRampStep 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "ramp"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Index           =   0
            Left            =   4005
            TabIndex        =   64
            Top             =   735
            Width           =   645
         End
      End
      Begin VB.CommandButton cmdPgUp 
         Caption         =   "Pg Next"
         DisabledPicture =   "frmPurgeProfile.frx":13D5C
         DownPicture     =   "frmPurgeProfile.frx":1445E
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   10005
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":14B60
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Next"
         DisabledPicture =   "frmPurgeProfile.frx":15262
         DownPicture     =   "frmPurgeProfile.frx":15964
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   9165
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":16066
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Prev"
         DisabledPicture =   "frmPurgeProfile.frx":16768
         DownPicture     =   "frmPurgeProfile.frx":16E6A
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":1756C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPgDn 
         Caption         =   "Pg Prev"
         DisabledPicture =   "frmPurgeProfile.frx":17C6E
         DownPicture     =   "frmPurgeProfile.frx":18370
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   6480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":18A72
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "Close"
         DisabledPicture =   "frmPurgeProfile.frx":19174
         DownPicture     =   "frmPurgeProfile.frx":19876
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   11220
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":19F78
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Close this screen"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmPurgeProfile.frx":1A67A
         DownPicture     =   "frmPurgeProfile.frx":1B2BC
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":1BEFE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Print a Listing of all Profiles"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore"
         DisabledPicture =   "frmPurgeProfile.frx":1CB40
         DownPicture     =   "frmPurgeProfile.frx":1D242
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":1D944
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Reload Station Recipe Values"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmPurgeProfile.frx":1E046
         DownPicture     =   "frmPurgeProfile.frx":1E748
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":1EE4A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Save Profile"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         DisabledPicture =   "frmPurgeProfile.frx":1F54C
         DownPicture     =   "frmPurgeProfile.frx":1FC4E
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmPurgeProfile.frx":20350
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Open Master Profile List"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Frame frmProfile 
         Caption         =   "Profile"
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
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1800
         Width           =   11955
         Begin VB.TextBox txtLastStepNumber 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Text            =   "100"
            ToolTipText     =   "total duration in hh:mm:ss"
            Top             =   270
            Width           =   1770
         End
         Begin VB.CommandButton cmdWizard 
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
            Left            =   11160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmPurgeProfile.frx":20A52
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Injection Information & Calculator"
            Top             =   240
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox txtProjectedLiters 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "12.23765"
            ToolTipText     =   "Projected Total Liters for this Profile"
            Top             =   270
            Width           =   1200
         End
         Begin VB.TextBox txtProjectedVolumes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Text            =   "1.221123"
            ToolTipText     =   "Projected Total Canister Volumes for this Profile"
            Top             =   570
            Width           =   1200
         End
         Begin VB.TextBox txtProfileDuration 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   2
            TabStop         =   0   'False
            Text            =   "hhh:mm:ss"
            ToolTipText     =   "Total duration as hhh:mm:ss"
            Top             =   570
            Width           =   1770
         End
         Begin VB.Label lblProjectedLiters 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   9480
            TabIndex        =   9
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label lblProjectedVolumes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "volumes"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   300
            Left            =   9480
            TabIndex        =   8
            Top             =   570
            Width           =   1095
         End
         Begin VB.Label lblProjected 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Projected"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   330
            Left            =   6810
            TabIndex        =   7
            Top             =   390
            Width           =   1200
         End
         Begin VB.Label lblLastStepNumber 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Steps"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   270
            Left            =   1080
            TabIndex        =   4
            Top             =   270
            Width           =   1290
         End
         Begin VB.Label lblProfileDuration 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Duration"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808000&
            Height          =   300
            Left            =   1080
            TabIndex        =   3
            Top             =   570
            Width           =   1290
         End
      End
      Begin Threed.SSPanel pnlDispProfNum 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   840
         Left            =   8160
         TabIndex        =   22
         ToolTipText     =   "Click for list of Defined Profiles"
         Top             =   120
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "49"
         ForeColor       =   -2147483646
         BackColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   1
         Font3D          =   3
      End
      Begin Threed.SSPanel pnlProfile 
         Height          =   510
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   11955
         _Version        =   65536
         _ExtentX        =   21087
         _ExtentY        =   900
         _StockProps     =   15
         Caption         =   "  Description: "
         ForeColor       =   -2147483646
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Begin VB.TextBox txtProfileDesc 
            BackColor       =   &H80000009&
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
            Height          =   285
            Left            =   1860
            MaxLength       =   70
            TabIndex        =   24
            ToolTipText     =   "Alphanumeric Description "
            Top             =   120
            Width           =   9915
         End
      End
      Begin VB.Label lblStnDesc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "station shift"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   4440
         TabIndex        =   184
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmPurgeProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ErrorModule 303     frmPurgeProfile
Option Explicit

'Dim errors As Integer
'Dim CurrUp, LastUp As Double
'Dim CurrDn, LastDn As Double
'Dim ChgValue As Integer

Private ProfileMode As Integer            ' 0=master; 1=station
Private DispProf As Integer               ' Current Master Profile index
Private ScreenBkgdColor As Long
Private ScreenDescription As String
Private ScreenDispFlag As Boolean
Private StnShftDescription As String
Private UpdatingAllSteps As Boolean
Private Chgs As Boolean
Private DspProfile As PurgeProfileType
Private MemProfile As PurgeProfileType
Private Idx As Integer
Private dbDbase As Database
Private rsProfile  As Recordset
Private rsSteps  As Recordset
Private Criteria As String
Private sdate As Date
Private sSec As Double
Private tmpval1, tmpval2, tmpval3 As Single
Private tmpStr As String
Private firstStep As Integer
Private dspStep As Integer
Private iStep As Integer
Private Mem_StepStartSetpoint(0 To 10) As Single
Private Mem_StepDuration(0 To 10) As Single
Private Mem_StepType(0 To 10) As Integer               ' 0 = undefined; 1 = step MfcSetPoint; 2 = ramp MfcSetPoint; 3 = last Step

Public Sub ChgProfileMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    ProfileMode = IIf((NewMode = MASTERMODE Or NewMode = STATIONMODE), NewMode, MASTERMODE)
    Select Case ProfileMode
        Case MASTERMODE
            ' station/shift description
            lblStnDesc.Visible = False
            ' screen description
            ScreenDescription = "Master Purge Profiles"
            ' screen background color
            ScreenBkgdColor = MasterMode_BackColor
            ' show MasterOnly items
            ScreenDispFlag = True
        Case STATIONMODE
            ' station/shift description
            StnShftDescription = "Station #" & Format(DispStn, "#0")
            If NR_SHIFT > 1 Then StnShftDescription = StnShftDescription & "  Shift #" & Format(DispShift, "#0")
            StnShftDescription = StnShftDescription & "  Purge Profile"
            lblStnDesc.Visible = True
            lblStnDesc.Left = cmdPgDn.Left
            lblStnDesc.ForeColor = TitlesData_Forecolor
            lblStnDesc.Caption = StnShftDescription
            ' screen description
            ScreenDescription = StnShftDescription
            ' screen background color
            ScreenBkgdColor = StationMode_BackColor
            ' hide MasterOnly items
            ScreenDispFlag = False
    End Select
    ' screen description
    frmPurgeProfile.Caption = ScreenDescription
    ' set screen background color
    frmPage.BackColor = ScreenBkgdColor
    pnlDispProfNum.BackColor = ScreenBkgdColor
    ' show Recipe # & Arrows ??
    cmdDown.Visible = ScreenDispFlag
    cmdUp.Visible = ScreenDispFlag
    cmdPgDn.Visible = ScreenDispFlag
    cmdPgUp.Visible = ScreenDispFlag
    pnlDispProfNum.Visible = ScreenDispFlag
End Sub

Public Sub LoadNewProf(ByVal NewProf As Integer)
    DispProf = NewProf
    ProfileDisplay_ByNum
End Sub

Private Function CalculatedLiters() As Single
' calculates the projected liters for the current PurgeProfile
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 123
Dim totalLiters As Single
Dim deltaLiters As Single
Dim deltaSP As Single
Dim avgSP As Single

    totalLiters = 0
    deltaLiters = 0
    For iStep = 1 To (MAX_PROFILESTEPS - 1)
        ' calculate liters for the current step
        Select Case DspProfile.StepType(iStep)
            Case STEPSTEP
                ' maintain this step's initial SP
                deltaLiters = DspProfile.StepStartSetpoint(iStep) * DspProfile.StepDuration(iStep)          ' sp in slpm; duration in min
            Case STEPRAMP
                ' ramp to next step's initial SP
                deltaSP = DspProfile.StepStartSetpoint(iStep + 1) - DspProfile.StepStartSetpoint(iStep)     ' sp in slpm
                avgSP = DspProfile.StepStartSetpoint(iStep) + (deltaSP / CSng(2))                           ' sp in slpm
                deltaLiters = avgSP * DspProfile.StepDuration(iStep)                                        ' sp in slpm; duration in min
            Case STEPLAST
                ' last step is always zero duration
                deltaLiters = 0
            Case Else
                ' undefined step
                deltaLiters = 0
        End Select
        totalLiters = totalLiters + deltaLiters
    Next iStep
        
    ' Projected Liters
    CalculatedLiters = totalLiters
    
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

Private Function CalculatedVolumes() As Single
' calculates the projected canister volumes for the current PurgeProfile
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 1234
Dim totalVolumes As Single

    totalVolumes = 0
    If ProfileMode = STATIONMODE Then
        If (StationCanister(DispStn, DispShift).WorkingVolume > 0) Then
            totalVolumes = DspProfile.ProjectedLiters / StationCanister(DispStn, DispShift).WorkingVolume
        End If
    End If
        
    ' Projected Volumes
    CalculatedVolumes = totalVolumes
    
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

Private Function CalculatedDuration() As Single
' calculates the projected duration (in minutes) for the current PurgeProfile
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 12345
Dim totalMinutes As Single
Dim deltaMinutes As Single

    totalMinutes = 0
    deltaMinutes = 0
    For iStep = 1 To MAX_PROFILESTEPS
        ' duration in minutes for the current step
        Select Case DspProfile.StepType(iStep)
            Case STEPSTEP
                ' maintain this step's initial SP
                deltaMinutes = DspProfile.StepDuration(iStep)    ' duration in min
            Case STEPRAMP
                ' ramp to next step's initial SP
                deltaMinutes = DspProfile.StepDuration(iStep)    ' duration in min
            Case STEPLAST
                ' last step is always zero duration
                deltaMinutes = 0
            Case Else
                ' undefined step
                deltaMinutes = 0
        End Select
        totalMinutes = totalMinutes + deltaMinutes
    Next iStep
        
    ' Projected Duration
    CalculatedDuration = totalMinutes
    
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

Private Function lastStepNumber() As Integer
Dim tempStep As Integer
    ' find highest "last Step"
    tempStep = 0
    For iStep = 1 To MAX_PROFILESTEPS
        ' check Initial SP & Duration for non-blank
        If DspProfile.StepType(iStep) = STEPLAST Then
            tempStep = iStep
        End If
    Next iStep
    ' Last Step Number
    lastStepNumber = tempStep
End Function

Private Sub cmdClear_Click(Index As Integer)
' clear block-of-10 boxes on screen
Dim idx2 As Integer
Dim iStep2 As Integer
    For iStep2 = 1 To 10
        idx2 = iStep2 + (CInt(10) * Index)
        txtStepNumber(idx2).text = Format((idx2 + firstStep - 1), "##0")
        txtInitialSP(idx2).text = "0.00"
        txtStepDuration(idx2).text = "0.000"
        chkStepStep(idx2).Value = 0
        chkRampStep(idx2).Value = 0
        chkLastStep(idx2).Value = 0
    Next iStep2
End Sub

Private Sub cmdCopy_Click()
    ScreenToDspProf
    DspProfToMemProf
End Sub

Private Sub cmdCopy10Steps_Click(Index As Integer)
' copy block-of-10 boxes on screen
Dim idx2 As Integer
Dim iStep2 As Integer
    For iStep2 = 1 To 10
        idx2 = iStep2 + (CInt(10) * Index)
        Mem_StepStartSetpoint(iStep2) = ValueFromText(txtInitialSP(idx2).text)
        Mem_StepDuration(iStep2) = ValueFromText(txtStepDuration(idx2).text)
        ' 0 = undefined; 1 = step MfcSetPoint; 2 = ramp MfcSetPoint; 3 = last Step
        If (chkStepStep(idx2).Value = cYES) Then
            Mem_StepType(iStep2) = STEPSTEP
        ElseIf (chkRampStep(idx2).Value = cYES) Then
            Mem_StepType(iStep2) = STEPRAMP
        ElseIf (chkLastStep(idx2).Value = cYES) Then
            Mem_StepType(iStep2) = STEPLAST
        Else
            Mem_StepType(iStep2) = NOSTEP
        End If
    Next iStep2
End Sub

Private Sub cmdDown_Click()
  DispProf = IIf(DispProf < 2, NR_PRGPROF, DispProf - 1)
  ProfileDisplay_ByNum
End Sub

Private Sub cmdImport_Click()
    Import_PurgeProfile
    DspProfToScreen
End Sub

Private Sub cmdOpen_Click()
    frmSearchProf.Show
    frmSearchProf.ChgSelectionDestination CInt(profdestProfile)
End Sub

Private Sub cmdPaste_Click()
    MemProfToDspProf
    DspProfile.Number = DispProf
    DspProfToScreen
    Chgs = True
End Sub

Private Sub cmdPaste10Steps_Click(Index As Integer)
' paste to block-of-10 boxes on screen
Dim idx2 As Integer
Dim iStep2 As Integer
    For iStep2 = 1 To 10
        idx2 = iStep2 + (CInt(10) * Index)
        txtStepNumber(idx2) = Format((idx2 + firstStep - 1), "##0")
        txtInitialSP(idx2).text = Format(Mem_StepStartSetpoint(iStep2), "##0.00")
        txtStepDuration(idx2).text = Format(Mem_StepDuration(iStep2), "#,##0.000")
        ' 0 = undefined; 1 = step MfcSetPoint; 2 = ramp MfcSetPoint; 3 = last Step
        chkStepStep(idx2).Value = IIf((Mem_StepType(iStep2) = STEPSTEP), cYES, cNO)
        chkRampStep(idx2).Value = IIf((Mem_StepType(iStep2) = STEPRAMP), cYES, cNO)
        chkLastStep(idx2).Value = IIf((Mem_StepType(iStep2) = STEPLAST), cYES, cNO)
    Next iStep2
End Sub

Private Sub cmdUp_Click()
  DispProf = IIf(DispProf > NR_PRGPROF - 1, 1, DispProf + 1)
  ProfileDisplay_ByNum
End Sub

Private Sub cmdPgDn_Click()
  DispProf = IIf(DispProf < 11, NR_PRGPROF, DispProf - 10)
  ProfileDisplay_ByNum
End Sub

Private Sub cmdPgUp_Click()
  DispProf = IIf(DispProf > NR_PRGPROF - 10, 1, DispProf + 10)
  ProfileDisplay_ByNum
End Sub

Private Sub cmdPrint_Click()
'    Print_TempVersTime 99
End Sub

Private Sub cmdRestore_Click()
    DspProfile = MemProfile
'    Profile-To-Screen
End Sub

Private Sub cmdStepPageDn_Click()
    firstStep = IIf((firstStep > 10), (firstStep - 10), 1)
    cmdStepPageDn.Enabled = IIf((firstStep <= 1), False, True)
    cmdStepPageUp.Enabled = IIf((firstStep >= (MAX_PROFILESTEPS - 19)), False, True)
    DspProfToScreen
End Sub

Private Sub cmdStepPageUp_Click()
    firstStep = IIf((firstStep < (MAX_PROFILESTEPS - 20)), (firstStep + 10), (MAX_PROFILESTEPS - 19))
    cmdStepPageDn.Enabled = IIf((firstStep <= 1), False, True)
    cmdStepPageUp.Enabled = IIf((firstStep >= (MAX_PROFILESTEPS - 19)), False, True)
    DspProfToScreen
End Sub

Private Sub cmdWizard_Click()
'    show Wizard screen
End Sub

Private Sub Form_Load()
Dim tmpColor As Long
Dim Idx As Integer

    KeyPreview = True
    
    ' Set Title Foreground color
    tmpColor = TitlesData_Forecolor
        frmProfile.ForeColor = tmpColor
        frmSteps.ForeColor = tmpColor
        pnlProfile.ForeColor = tmpColor
    tmpColor = Titles_ForeColor
        lblLastStepNumber.ForeColor = tmpColor
        lblProfileDuration.ForeColor = tmpColor
        lblProjected.ForeColor = tmpColor
        For Idx = 0 To 1
            lblStep(Idx).ForeColor = tmpColor
            lblStepUnits(Idx).ForeColor = tmpColor
            lblSpDesc(Idx).ForeColor = tmpColor
            lblSP(Idx).ForeColor = tmpColor
            lblSpUnits(Idx).ForeColor = tmpColor
            lblDurationDesc(Idx).ForeColor = tmpColor
            lblDuration(Idx).ForeColor = tmpColor
            lblDurationUnits(Idx).ForeColor = tmpColor
            lblStepTypeDesc(Idx).ForeColor = tmpColor
            lblStepType(Idx).ForeColor = tmpColor
            lblStepStep(Idx).ForeColor = tmpColor
            lblRampStep(Idx).ForeColor = tmpColor
            lblLastStep(Idx).ForeColor = tmpColor
        Next Idx
    tmpColor = TitlesLabel_ForeColor
        txtLastStepNumber.ForeColor = tmpColor
        txtProfileDuration.ForeColor = tmpColor
        txtProjectedLiters.ForeColor = tmpColor
        lblProjectedLiters.ForeColor = tmpColor
        txtProjectedVolumes.ForeColor = tmpColor
        lblProjectedVolumes.ForeColor = tmpColor
        
    ' Reset all the backgrounds
'    Reset_BackColors

    lblMessage.Caption = ""

    
    cmdPrint.Enabled = IIf(PRINTERAVAILABLE, True, False)
    cmdStepPageDn.Enabled = False
    cmdStepPageUp.Enabled = True
    firstStep = 1
    DispProf = 0
    pnlDispProfNum.Caption = "0"
    txtProfileDesc.text = " "
    UpdatingAllSteps = False

    ' authorized to Save Master Profiles ?
    cmdSave.Visible = IIf(CheckPass("O", False), True, False)
    
    ' show Restore Station Profile ?
    cmdRestore.Visible = IIf(ProfileMode = STATIONMODE, True, False)
        
    ' open profile / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    ExitScreen
End Sub

Private Sub txtInitialSP_LostFocus(Index As Integer)
Dim PurgeMfcMaxSlpm As Single
    txtInitialSP(Index).BackColor = vbWhite
    If (Not UpdatingAllSteps) Then
        If (IsNumeric(txtInitialSP(Index).text)) Then
            Select Case ProfileMode
                Case MASTERMODE
                    PurgeMfcMaxSlpm = 1000
                Case STATIONMODE
                    PurgeMfcMaxSlpm = Stn_AIO(DispStn, asPurgeAirFlow).EuMax
            End Select
            If Range_Check(txtInitialSP(Index), 0, PurgeMfcMaxSlpm, "Initial Flow SP") Then
                lblMessage.Caption = ""
                DspProfile.StepStartSetpoint(firstStep + Index - 1) = ValueFromText(txtInitialSP(Index).text)
                DspProfToScreen
            End If
        End If
    End If
End Sub

Private Sub txtStepDuration_LostFocus(Index As Integer)
    txtStepDuration(Index).BackColor = vbWhite
    If (Not UpdatingAllSteps) Then
        If (IsNumeric(txtStepDuration(Index).text)) Then
            If Range_Check(txtStepDuration(Index), 0, 1000, "Step Duration") Then
                lblMessage.Caption = ""
                DspProfile.StepDuration(firstStep + Index - 1) = ValueFromText(txtStepDuration(Index).text)
                DspProfToScreen
            End If
        End If
    End If
End Sub

Private Sub SetDspProfStepType(ByVal idxChkBox As Integer, ByVal idxStep As Integer)
    If (chkLastStep(idxChkBox).Value = cYES) Then
        DspProfile.StepType(idxStep) = STEPLAST
    ElseIf (chkStepStep(idxChkBox).Value = cYES) Then
        DspProfile.StepType(idxStep) = STEPSTEP
    ElseIf (chkRampStep(idxChkBox).Value = cYES) Then
        DspProfile.StepType(idxStep) = STEPRAMP
    Else
        DspProfile.StepType(idxStep) = NOSTEP
    End If
End Sub

Private Sub chkLastStep_Click(Index As Integer)
    chkLastStep(Index).BackColor = vbWhite
    If (Not UpdatingAllSteps) Then
        If (chkLastStep(Index).Value = cYES) Then
            DspProfile.StepType(firstStep + Index - 1) = STEPLAST
            DspProfile.StepDuration(firstStep + Index - 1) = 0
        Else
            SetDspProfStepType Index, (firstStep + Index - 1)
        End If
        DspProfToScreen
    End If
End Sub

Private Sub chkRampStep_Click(Index As Integer)
    chkRampStep(Index).BackColor = vbWhite
    If (Not UpdatingAllSteps) Then
        If (chkRampStep(Index).Value = cYES) Then
            DspProfile.StepType(firstStep + Index - 1) = STEPRAMP
        Else
            SetDspProfStepType Index, (firstStep + Index - 1)
        End If
        DspProfToScreen
    End If
End Sub

Private Sub chkStepStep_Click(Index As Integer)
    chkStepStep(Index).BackColor = vbWhite
    If (Not UpdatingAllSteps) Then
        If (chkStepStep(Index).Value = cYES) Then
            DspProfile.StepType(firstStep + Index - 1) = STEPSTEP
        Else
            SetDspProfStepType Index, (firstStep + Index - 1)
        End If
        DspProfToScreen
    End If
End Sub

Private Sub txtProfileDesc_Change()
    DspProfile.Description = Trim(txtProfileDesc.text)
    txtProfileDesc.BackColor = vbWhite
End Sub



'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************
'**************************************************************



Public Function OkToRunProfileInStation() As Boolean
    OkToRunProfileInStation = ValidProfile
End Function

Public Sub ExportProfile()
    ScreenToDspProf
    ExportedProfile = DspProfile
End Sub

Private Sub DspProfToScreen()
' Copies DspProfile to Screen
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 9
Dim flag As Boolean
Dim Idx As Integer

    UpdatingAllSteps = True
    
    ' update cumulative values
    DspProfile.EndStep = lastStepNumber
    DspProfile.Duration = CalculatedDuration
    DspProfile.DurDesc = ProfileDurationDescription(DspProfile.Duration)
    DspProfile.ProjectedLiters = CalculatedLiters
    DspProfile.ProjectedVolumes = CalculatedVolumes
    
    ' description
    txtProfileDesc.text = DspProfile.Description
    
    ' Profile #, Arrows & overall values
    Select Case ProfileMode
        Case MASTERMODE
            ' Profile # & Arrows
            cmdDown.Visible = True
            cmdUp.Visible = True
            cmdPgDn.Visible = True
            cmdPgUp.Visible = True
            pnlDispProfNum.Visible = True
            pnlDispProfNum.Caption = Format(DspProfile.Number, "#00")
            ' overall values
            txtLastStepNumber.text = Format(DspProfile.EndStep, "##0")
            txtProfileDuration.text = DspProfile.DurDesc
            lblProjectedLiters.Top = 390
            txtProjectedLiters.Top = 390
            txtProjectedLiters.text = Format(DspProfile.ProjectedLiters, "#,###,##0.00")
            lblProjectedVolumes.Top = OutOfSight
            txtProjectedVolumes.Top = OutOfSight
        Case STATIONMODE
            ' Profile # & Arrows
            cmdDown.Visible = False
            cmdUp.Visible = False
            cmdPgDn.Visible = False
            cmdPgUp.Visible = False
            pnlDispProfNum.Visible = False
            ' overall values
            txtLastStepNumber.text = Format(DspProfile.EndStep, "##0")
            txtProfileDuration.text = DspProfile.DurDesc
            lblProjectedLiters.Top = 270
            txtProjectedLiters.Top = 270
            txtProjectedLiters.text = Format(DspProfile.ProjectedLiters, "#,###,##0.00")
            lblProjectedVolumes.Top = 570
            txtProjectedVolumes.Top = 570
            txtProjectedVolumes.text = Format(DspProfile.ProjectedVolumes, "###,##0.000")
        Case Else
            Exit Sub
    End Select
    
    ' displayed steps
    For Idx = 1 To 20
        iStep = firstStep + Idx - 1
        txtStepNumber(Idx).text = Format(iStep, "##0")
        txtInitialSP(Idx).text = Format(DspProfile.StepStartSetpoint(iStep), "##0.00")
        txtStepDuration(Idx).text = Format(DspProfile.StepDuration(iStep), "#,##0.000")
        chkStepStep(Idx).Value = IIf((DspProfile.StepType(iStep) = STEPSTEP), cYES, cNO)
        chkRampStep(Idx).Value = IIf((DspProfile.StepType(iStep) = STEPRAMP), cYES, cNO)
        chkLastStep(Idx).Value = IIf((DspProfile.StepType(iStep) = STEPLAST), cYES, cNO)
    Next Idx

    UpdatingAllSteps = False
    
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

Private Sub DspProfToMemProf()
    MemProfile = DspProfile
End Sub

Private Sub MemProfToDspProf()
    DspProfile = MemProfile
End Sub

Private Sub ScreenToDspProf()
' Copies Screen data to DspProfile
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 5
Dim Idx As Integer

    DspProfile.Number = 0
    If IsNumeric(pnlDispProfNum.Caption) Then DspProfile.Number = CInt(pnlDispProfNum.Caption)
    DspProfile.Description = Trim(txtProfileDesc.text)
    
    ' displayed steps
    For Idx = 1 To 20
        iStep = firstStep + Idx - 1
        DspProfile.StepStartSetpoint(iStep) = ValueFromText(txtInitialSP(Idx).text)
        DspProfile.StepDuration(iStep) = ValueFromText(txtStepDuration(Idx).text)
        If (chkStepStep(Idx).Value = cYES) Then
            DspProfile.StepType(iStep) = STEPSTEP
        ElseIf (chkRampStep(Idx).Value = cYES) Then
            DspProfile.StepType(iStep) = STEPRAMP
        ElseIf (chkLastStep(Idx).Value = cYES) Then
            DspProfile.StepType(iStep) = STEPLAST
        Else
            DspProfile.StepType(iStep) = NOSTEP
        End If
    Next Idx

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

Private Sub ExitScreen()
    ' close profile database
    dbDbase.Close
    ' unload form
'    frmPurgeProfile.Visible = False
    Set frmPurgeProfile = Nothing
    Unload Me
End Sub

Public Sub ProfileDisplay_ByNum()
    GetProfile MASTERMODE, DispProf, 0
    DspProfToScreen
    Chgs = False
End Sub

Public Sub ProfileDisplay_ByStnShift()
    GetProfile STATIONMODE, DispStn, DispShift
    DspProfToScreen
    Chgs = False
End Sub

Private Sub GetProfile(ByVal MstStnMode As Integer, ByVal idx1 As Integer, ByVal idx2 As Integer)
    Select Case MstStnMode
        Case MASTERMODE
            ' Read Master Profile Record
            Criteria = "SELECT * FROM [MasterProfiles] WHERE [Number] = " & idx1 & " "
            Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            ' Read Master Profile Steps
            Criteria = "SELECT * FROM [MasterProfileSteps] WHERE [ProfileNumber] = " & idx1 & " ORDER BY [StepNumber] ASC"
            Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
        Case STATIONMODE
            ' Read Station Profile Record
            Criteria = "SELECT * FROM [StationProfiles] WHERE [Station] = " & idx1 & "  and [Shift] = " & idx2 & " "
            Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            ' Read Station Profile Steps
            Criteria = "SELECT * FROM [StationProfileSteps] WHERE [Station] = " & idx1 & "  and [Shift] = " & idx2 & " ORDER BY [StepNumber] ASC"
            Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
    End Select
    If rsProfile.BOF Then
        InitDspProf MstStnMode, idx1, idx2
    Else
        DbToDspProf
    End If
    rsProfile.Close
    rsSteps.Close
End Sub

Private Sub DbToDspProf()
'
Dim Idx As Integer
    ' Load Profile Record to DspProfile
    DispProf = rsProfile("Number")
    DspProfile.Number = rsProfile("Number")
    DspProfile.Description = rsProfile("Description")
   
    DspProfile.Duration = rsProfile("TotalDuration")
    DspProfile.DurDesc = ProfileDurationDescription(DspProfile.Duration)
    DspProfile.EndStep = rsProfile("Steps")
    DspProfile.ProjectedLiters = rsProfile("ProjectedLiters")
    DspProfile.ProjectedVolumes = rsProfile("ProjectedVolumes")
    
    ' steps
    For Idx = 1 To MAX_PROFILESTEPS
        DspProfile.StepDuration(Idx) = 0
        DspProfile.StepStartSetpoint(Idx) = 0
        DspProfile.StepType(Idx) = 0
    Next Idx
    If (Not rsSteps.BOF) Then
        rsSteps.MoveFirst
        While Not rsSteps.EOF
            Idx = rsSteps("StepNumber")
            DspProfile.StepDuration(Idx) = rsSteps("Duration")
            DspProfile.StepStartSetpoint(Idx) = rsSteps("InitialSP")
            DspProfile.StepType(Idx) = rsSteps("StepType")
            rsSteps.MoveNext
        Wend
    End If
                
End Sub

Private Sub SaveMasterProf(ByVal iProf As Integer)
'
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 111
        
    ' Save Master PurgeProfile Record
    Criteria = "SELECT * FROM [MasterProfiles] WHERE [Number] = " & iProf & " "
    Set rsProfile = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    If rsProfile.BOF Then
        rsProfile.AddNew
        rsProfile("Number") = iProf
    Else
      rsProfile.MoveFirst
      rsProfile.Edit
    End If
       
    ' Update Master PurgeProfile Record
    rsProfile("Description") = DspProfile.Description
    rsProfile("TotalDuration") = DspProfile.Duration
    rsProfile("Steps") = DspProfile.EndStep
    rsProfile("ProjectedLiters") = DspProfile.ProjectedLiters
    rsProfile("ProjectedVolumes") = DspProfile.ProjectedVolumes
    rsProfile.Update
    rsProfile.Close

    ' Save Master PurgeProfile Steps
    Criteria = "SELECT * FROM [MasterProfileSteps] WHERE [ProfileNumber] = " & iProf & " ORDER BY [StepNumber] ASC"
    Set rsSteps = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    ' first remove existing steps
    If rsSteps.BOF Then
        ' nothing to do; no steps for this profile exist in db
    Else
        rsSteps.MoveLast
        If (Not rsSteps.BOF) Then
            While Not rsSteps.BOF
                rsSteps.Delete
                rsSteps.MovePrevious
            Wend
        End If
    End If
       
    ' Update Master PurgeProfile Steps
'    rsSteps.MoveFirst
    For iStep = 1 To MAX_PROFILESTEPS
        If DspProfile.StepType(iStep) <> NOSTEP Then
            rsSteps.AddNew
            rsSteps("ProfileNumber") = iProf
            rsSteps("StepNumber") = iStep
            rsSteps("Duration") = DspProfile.StepDuration(iStep)
            rsSteps("InitialSP") = DspProfile.StepStartSetpoint(iStep)
            rsSteps("StepType") = DspProfile.StepType(iStep)
            rsSteps("StepTypeDesc") = PurgeProfileStepDesc(DspProfile.StepType(iStep))
            rsSteps.Update
        End If
    Next iStep
    rsSteps.Close

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

Public Sub InitDspProf(ByVal MstStnMode As Integer, ByVal idx1 As Integer, ByVal idx2 As Integer)
' Initializes DspProfile
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 1
Dim Idx As Integer

    Select Case MstStnMode
        Case MASTERMODE
            ' master
            DspProfile.Description = "undefined"
            DspProfile.Number = CInt(idx1)
            
            DspProfile.Duration = CSng(0)
            DspProfile.DurDesc = ProfileDurationDescription(DspProfile.Duration)
            DspProfile.EndStep = CInt(0)
            DspProfile.ProjectedLiters = CSng(0)
            DspProfile.ProjectedVolumes = CSng(0)
        
            ' steps
            For Idx = 1 To MAX_PROFILESTEPS
                DspProfile.StepDuration(Idx) = CSng(0)
                DspProfile.StepStartSetpoint(Idx) = CSng(0)
                DspProfile.StepType(Idx) = CInt(0)
            Next Idx
    
        Case STATIONMODE
            ' station
            DspProfile.Description = "undefined"
            DspProfile.Number = CInt(0)
            
            DspProfile.Duration = CSng(0)
            DspProfile.DurDesc = ProfileDurationDescription(DspProfile.Duration)
            DspProfile.EndStep = 0
            DspProfile.ProjectedLiters = CSng(0)
            DspProfile.ProjectedVolumes = CSng(0)
        
            ' steps
            For Idx = 1 To MAX_PROFILESTEPS
                DspProfile.StepDuration(Idx) = CSng(0)
                DspProfile.StepStartSetpoint(Idx) = CSng(0)
                DspProfile.StepType(Idx) = CInt(0)
            Next Idx
    
    End Select


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

Public Sub InitProfile()
    Select Case ProfileMode
        Case MASTERMODE
            ' master
            If (DispProf < 1 Or DispProf > NR_PRGPROF) Then DispProf = 1
            GetProfile MASTERMODE, DispProf, 0
            cmdRestore.Visible = False
            cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
        Case STATIONMODE
            ' station
            If StationProfile(DispStn, DispShift).Number < 0 _
             Or StationProfile(DispStn, DispShift).Number > NR_PRGPROF Then
               DispProf = 0
            Else
               DispProf = StationProfile(DispStn, DispShift).Number
            End If
            GetProfile STATIONMODE, DispStn, DispShift
            If StationControl(DispStn, DispShift).Mode <> VBIDLE Then
               cmdRestore.Visible = False
               cmdSave.Visible = False
            Else
               cmdRestore.Visible = True
               cmdSave.Visible = True
            End If
            cmdPrint.Visible = False
    End Select
    DspProfToScreen
    Chgs = False
End Sub

Private Sub UpdateProfile()
    If (DispProf < 1 Or DispProf > NR_PRGPROF) Then DispProf = 1
    GetProfile MASTERMODE, DispProf, 0
    DspProfToScreen
    Chgs = False
End Sub

Private Sub cmdSave_Click()
    SaveProfile
End Sub

Private Sub cmdReturn_Click()
    ExitScreen
End Sub

Private Function ValidProfile() As Boolean
' Function Name:    ValidProfile
' Description:      Checks the validity of Profile settings.
'                   Used before saving the Profile.
'                   Returns a true value if values are okay.
'                   Returns a false value if values are not okay.
'                   If an error is detected, an appropriate message
'                   is displayed.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 313
Dim Idx As Integer
Dim Message As String
Dim lastSteps As Integer
Dim lastStepNumber As Integer
Dim validFlag As Boolean

    validFlag = True
    lblMessage.Caption = ""
    
    ' Name
    If Len(txtProfileDesc.text) > 50 Then
        validFlag = False
        txtProfileDesc.BackColor = EntryInvalid_BackColor
        Message = "Name is too long. 50 char Max"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    ElseIf Len(txtProfileDesc.text) < 1 Then
        txtProfileDesc.text = " "
    End If
    
    ' Steps
    lastSteps = 0
    For iStep = 1 To MAX_PROFILESTEPS
        If (DspProfile.StepType(iStep) = STEPLAST) Then
            lastSteps = lastSteps + 1
            lastStepNumber = iStep
        End If
    Next iStep
    
    ' Last Step
    Select Case lastSteps
        Case 0
            ' no last step; not enough
            validFlag = False
            Message = "There is no designated Last Step"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        Case 1
            ' one last step; just right
        Case Else
            ' more than one last step; too much
            validFlag = False
            Message = "There is more than one designated Last Step"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            For iStep = 1 To MAX_PROFILESTEPS
                If (DspProfile.StepType(iStep) = STEPLAST) Then
                    If ((iStep >= firstStep) And (iStep <= (firstStep + 19))) Then
                        Idx = iStep - firstStep + 1
                        chkLastStep(Idx).BackColor = EntryInvalid_BackColor
                    End If
                End If
            Next iStep
    End Select
    
    ' ************************************************************************
    ' Check for undefined steps
    ' ************************************************************************
    If validFlag Then
        For iStep = 1 To MAX_PROFILESTEPS
            If (iStep <= lastStepNumber) Then
                If (DspProfile.StepType(iStep) = NOSTEP) Then
                    validFlag = False
                    Message = "Step #" & Format(iStep, "###0") & " is Undefined"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                    If ((iStep >= firstStep) And (iStep <= (firstStep + 19))) Then
                        Idx = iStep - firstStep + 1
                        chkStepStep(Idx).BackColor = EntryInvalid_BackColor
                        chkRampStep(Idx).BackColor = EntryInvalid_BackColor
                        chkLastStep(Idx).BackColor = EntryInvalid_BackColor
                    End If
                End If
            End If
        Next iStep
    End If

    ' ************************************************************************
    ' Cleanup values for unused steps
    ' ************************************************************************
    If validFlag Then
        For iStep = 1 To MAX_PROFILESTEPS
            If (iStep > lastStepNumber) Then
                DspProfile.StepStartSetpoint(iStep) = 0
                DspProfile.StepDuration(iStep) = 0
                DspProfile.StepType(iStep) = NOSTEP
            End If
        Next iStep
    End If

    ' ************************************************************************
    ' Additional Validation Checks when saving a Profile to a specific station
    ' ************************************************************************
    If ProfileMode = STATIONMODE Then
        If validFlag Then
            For iStep = 1 To MAX_PROFILESTEPS
                If (DspProfile.StepStartSetpoint(iStep) > Stn_AIO(DispStn, asPurgeAirFlow).EuMax) Then
                    validFlag = False
                    Message = "Initial SP cannot exceed Purge MFC Range"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                    If ((iStep >= firstStep) And (iStep <= (firstStep + 19))) Then
                        Idx = iStep - firstStep + 1
                        txtInitialSP(Idx).BackColor = EntryInvalid_BackColor
                    End If
                End If
                If (iStep < lastStepNumber) Then
                    If (DspProfile.StepDuration(iStep) = 0) Then
                        validFlag = False
                        Message = "Step Duration is zero"
                        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                        If ((iStep >= firstStep) And (iStep <= (firstStep + 19))) Then
                            Idx = iStep - firstStep + 1
                            txtStepDuration(Idx).BackColor = EntryInvalid_BackColor
                        End If
                    End If
                End If
            Next iStep
        End If
    End If

ValidProfile = validFlag

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

Private Function Range_Check(tcontrol As Control, ByVal slow As Single, ByVal shigh As Single, ByVal slabel As String) As Boolean
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
SetErrModule 303, 333
If UseLocalErrorHandler Then On Error GoTo localhandler
Dim svalue As Single
Dim Message As String
Dim flag As Boolean

    flag = True
    If (tcontrol.text = Empty) Then
        
        ' Empty Value
        flag = False
        tcontrol.BackColor = EntryInvalid_BackColor
        Message = slabel & ":  Value is Empty!"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    
    ElseIf Not IsNumeric(tcontrol.text) Then
        
        ' Non-Numeric Value
        flag = False
        tcontrol.BackColor = EntryInvalid_BackColor
        Message = slabel & ":  Value is Not Numeric!"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    
    Else
    
        ' Numeric Value
        svalue = CSng(tcontrol.text)
        
        ' Check Value against Limits
        If svalue < slow Or svalue > shigh Then
            flag = False
            tcontrol.BackColor = EntryInvalid_BackColor
        '    tcontrol.SelStart = 0
        '    tcontrol.SelLength = Len(tcontrol.text)
        '    tcontrol.SetFocus
            Message = slabel & ":  Value out of range! " & "( " & Format(slow, "###0.00") & " - " & Format(shigh, "###0.00") & " )"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        End If
    End If
    
    Range_Check = flag

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

Public Sub SaveProfile()
SetErrModule 303, 2
If UseLocalErrorHandler Then On Error GoTo localhandler

    Select Case ProfileMode
        Case MASTERMODE
            ' master
            If CheckPass("O", False) Then
                lblMessage.Caption = vbCrLf
                Reset_BackColors
                If ValidProfile Then
                    Reset_BackColors
                    ScreenToDspProf              ' Copy screen data to Profile Array
                    ' Save Master Profile Information
                    SaveMasterProf CInt(DspProfile.Number)
                    ' Save Remote Master Profile Information
                    If USINGREMCANLOAD Then
                        ' open master canister / recipe database
                        Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
                        ' open remote database
                        OpenConnToRemoteDb
                        ' update Remote Master Profile Information
                        UpdateRemotePurgeProfiles
                        ' close remote database
                        CloseConnToRemoteDb
                    End If
                    Chgs = False
                    lblMessage.ForeColor = Message_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "New Profile Settings Saved" & vbCrLf
                Else
                    lblMessage.ForeColor = Alarm_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "Profile Settings Not Saved" & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "Try again after correcting the errors." & vbCrLf
                    Beep
                    Beep
                    Beep
                End If
            Else
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Insufficient Access" & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "The current user is not authorized to save Master Profiles" & vbCrLf
                Beep
                Beep
                Beep
            End If
        
        Case STATIONMODE
            ' station
            If StationControl(DispStn, DispShift).Mode = VBIDLE Then
                lblMessage.Caption = vbCrLf
                Reset_BackColors
                If ValidProfile Then
                    Reset_BackColors
                    ' Copy screen data to Station Profile
                    ScreenToDspProf
                    ' Update Station Profile Description & (Master Profile)Number
                    StationProfile(DispStn, DispShift) = DspProfile
                    StationProfile(DispStn, DispShift).Number = IIf(Chgs, CInt(0), DispProf)
                    ' save station Profiles
                    Save_StationProfiles
                    lblMessage.ForeColor = Message_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    Select Case NR_SHIFT
                        Case 1
                            lblMessage.Caption = lblMessage.Caption & "New Profile Settings Saved to Station #" + Format(DispStn, "0")
                        Case 2, 3, 4
                            lblMessage.Caption = lblMessage.Caption & "New Profile Settings Saved to Station #" + Format(DispStn, "0") + " / Shift #" + Format(DispShift, "0")
                    End Select
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "Estimated Profile Duration is " + StationProfile(DispStn, DispShift).DurDesc
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                End If
            Else
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Profile Settings Not Saved" & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Can Not Change values while station is running" & vbCrLf
                Beep
                Beep
                Beep
            End If

    End Select

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

Private Sub Reset_BackColors()
'
' resets the background colors
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 6
Dim Idx As Integer

    txtProfileDesc.BackColor = txtNotHighlight.BackColor
    
    For Idx = 1 To 20
        txtStepNumber(Idx).BackColor = txtNotHighlight.BackColor
        txtInitialSP(Idx).BackColor = txtNotHighlight.BackColor
        txtStepDuration(Idx).BackColor = txtNotHighlight.BackColor
        chkStepStep(Idx).BackColor = txtNotHighlight.BackColor
        chkRampStep(Idx).BackColor = txtNotHighlight.BackColor
        chkLastStep(Idx).BackColor = txtNotHighlight.BackColor
    Next Idx
    
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

Public Sub CopyAllRemoteMasterProfiles()
'
'        Copy ALL Master PurgeProfile Information Records to Remote DB
'
    ' reset RemoteProfiles RecordSource
    frmRemotePrf.adoRemoteProfiles.RecordSource = "SELECT * FROM [MasterProfiles] ORDER BY [MasterProfiles].[Number] ASC"
    frmRemotePrf.adoRemoteProfiles.Refresh
    frmRemotePrf.dgRemoteProfiles.Refresh
                
    ' reset RemoteProfileSteps RecordSource
    frmRemotePrf.adoRemoteProfileSteps.RecordSource = "SELECT * FROM [MasterProfiles] ORDER BY [MasterProfileSteps].[StepNumber] ASC"
'    frmRemotePrf.adoRemoteProfileSteps.Refresh
    frmRemotePrf.dgRemoteProfileSteps.Refresh
                

    Unload frmRemotePrf

End Sub

Sub Import_PurgeProfile()
' Import_PurgeProfile SP Information
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 303, 1475
Dim filler As String
Dim iFileNumber As Integer
Dim sFileName As String
Dim iStation As Integer
Dim iStep As Integer
Dim maxStep As Integer
Dim valDesc As String
Dim valSP As Single
Dim valDur As Single
Dim valType As Integer

    sFileName = FILEPATH_rcp & "importprofile.prg"
    iFileNumber = FreeFile
    Open sFileName For Input As #iFileNumber
    
    
    ' clear all steps
    For iStep = 1 To MAX_PROFILESTEPS
        DspProfile.StepDuration(iStep) = 0
        DspProfile.StepStartSetpoint(iStep) = 0
        DspProfile.StepType(iStep) = 0
    Next iStep
    
    '  Purge Profile Steps
    Input #iFileNumber, valDesc
    iStep = 0
    '     Do Until ((EOF(iFileNumber)) Or (iStep > MAX_PROFILESTEPS))
    Do Until EOF(iFileNumber)
        If (EOF(iFileNumber)) Then
            lblMessage.ForeColor = Alarm_ForeColor
            lblMessage.Caption = vbCrLf & "End-Of-File" & vbCrLf
        Else
            iStep = iStep + 1
            valSP = 0
            valDur = 0
            Input #iFileNumber, valSP, valDur
            If (MAX_PROFILESTEPS >= iStep) Then
                If ((valDur <> 0) And (valSP <> 0)) Then
                    DspProfile.StepDuration(iStep) = valDur
                    DspProfile.StepStartSetpoint(iStep) = valSP
                    DspProfile.StepType(iStep) = STEPRAMP
                    maxStep = iStep
                End If
            Else
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = vbCrLf & "Too many steps (" & Format(iStep, "###,##0") & " in import file; extra steps ignored" & vbCrLf
            End If
        End If
    Loop
    
ChgErrModule 303, 1476

    If (maxStep > MAX_PROFILESTEPS) Then maxStep = MAX_PROFILESTEPS
    ' make the last step a Last Step
    DspProfile.StepType(maxStep) = STEPLAST
    
    ' Purge Profile Record to DspProfile
    DspProfile.Number = DispProf
    DspProfile.Description = IIf((Len(valDesc) > 1), valDesc, "imported")
    DspProfile.Duration = 0
    DspProfile.DurDesc = ProfileDurationDescription(DspProfile.Duration)
    DspProfile.EndStep = iStep
    DspProfile.ProjectedLiters = 0
    DspProfile.ProjectedVolumes = 0
    
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = vbCrLf & "Import of " & Format(maxStep, "###,##0") & "  steps completed" & vbCrLf
        
    Close #iFileNumber
        
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
If err.Number = 62 Then
    iresponse = vbIgnore
Else
    iresponse = ErrorHandler(err)
End If
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



