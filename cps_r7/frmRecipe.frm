VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRecipe 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recipe Details"
   ClientHeight    =   11325
   ClientLeft      =   900
   ClientTop       =   645
   ClientWidth     =   12840
   ClipControls    =   0   'False
   Icon            =   "frmRecipe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11325
   ScaleWidth      =   12840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmRcpInfo 
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
      Height          =   4575
      Left            =   4440
      TabIndex        =   160
      Top             =   10560
      Width           =   6615
      Begin VB.CommandButton cmdCloseInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   163
         ToolTipText     =   "Close the Recipe End Method Information "
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   3720
         Left            =   120
         Picture         =   "frmRecipe.frx":5CD4
         ScaleHeight     =   3660
         ScaleWidth      =   6285
         TabIndex        =   161
         Top             =   720
         Width           =   6345
      End
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stable Weight Change End Method Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   162
         Top             =   240
         Width           =   6105
      End
   End
   Begin VB.Frame frmCycleType 
      Caption         =   "Cycles"
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
      Height          =   855
      Left            =   120
      TabIndex        =   157
      Top             =   6495
      Width           =   4095
      Begin MSComctlLib.TabStrip tabsCycletype 
         Height          =   510
         Left            =   60
         TabIndex        =   158
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   900
         TabWidthStyle   =   2
         Style           =   1
         TabFixedWidth   =   3403
         TabFixedHeight  =   900
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Purge - Load"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Load - Purge"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   14400
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   81
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Frame frmStart 
      Caption         =   "Start Method"
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
      Height          =   1275
      Left            =   120
      TabIndex        =   151
      Top             =   5195
      Width           =   4100
      Begin VB.TextBox txtStartAfterMin 
         Alignment       =   2  'Center
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
         Left            =   2000
         MaxLength       =   3
         TabIndex        =   156
         Text            =   "0"
         ToolTipText     =   "1 to 999 Minutes"
         Top             =   585
         Width           =   715
      End
      Begin VB.CheckBox optStartAfter 
         Caption         =   "Delay Start by                        minutes"
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
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   155
         Top             =   600
         Width           =   3765
      End
      Begin VB.TextBox txtStartAtDate 
         Alignment       =   2  'Center
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
         Left            =   1320
         TabIndex        =   154
         Text            =   "DD/MM/YYYY hh:mm"
         ToolTipText     =   "Enter Start Time as DD/MM/YYYY hh:mm"
         Top             =   885
         Width           =   2385
      End
      Begin VB.CheckBox optStartNow 
         Caption         =   "Start without delay"
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
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   153
         Top             =   300
         Width           =   3645
      End
      Begin VB.CheckBox optStartAt 
         Caption         =   "Start at "
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
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   152
         Top             =   900
         Width           =   3645
      End
   End
   Begin VB.Frame frmEnd 
      Caption         =   "End Method"
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
      Height          =   2300
      Left            =   120
      TabIndex        =   137
      Top             =   1570
      Width           =   4100
      Begin VB.CheckBox optUpdateCanWc 
         Alignment       =   1  'Right Justify
         Caption         =   "Update Canister Working Capacity"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   164
         ToolTipText     =   "Update Canister Working Capacity at the end of a successfull test?"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.CommandButton cmdEndMethodInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":5054E
         Style           =   1  'Graphical
         TabIndex        =   159
         ToolTipText     =   "Information about the End Method Choices"
         Top             =   1560
         UseMaskColor    =   -1  'True
         Width           =   715
      End
      Begin VB.TextBox txtPFCycle 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   139
         Text            =   "999"
         ToolTipText     =   "1 to 999"
         Top             =   308
         Width           =   570
      End
      Begin VB.CheckBox optEndWeightChange 
         Caption         =   "End after Stable Weight Chg Per Load"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   144
         Top             =   660
         Width           =   3765
      End
      Begin VB.TextBox txtConsecutiveCycles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2000
         MaxLength       =   3
         TabIndex        =   143
         Text            =   "999"
         ToolTipText     =   "Number of consecutive cycles with weight change within tolerance"
         Top             =   1275
         Width           =   450
      End
      Begin VB.TextBox txtMinimumCycles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3435
         MaxLength       =   3
         TabIndex        =   142
         Text            =   "0"
         ToolTipText     =   "Minimum number of cycles "
         Top             =   1275
         Width           =   450
      End
      Begin VB.TextBox txtWeightChangeTol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2550
         MaxLength       =   5
         TabIndex        =   141
         Text            =   "0.001"
         ToolTipText     =   "Weight change tolerance in grams"
         Top             =   975
         Width           =   630
      End
      Begin VB.CheckBox optEndCycles 
         Caption         =   "End after             Purge / Load Cycles"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   140
         Top             =   300
         Width           =   3765
      End
      Begin VB.TextBox txtMaximumCycles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3435
         MaxLength       =   3
         TabIndex        =   138
         Text            =   "0"
         ToolTipText     =   "Maximum number of cycles "
         Top             =   1530
         Width           =   450
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   150
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblConsecutiveCycles 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cycles:  Consecutive"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   149
         Top             =   1290
         Width           =   1845
      End
      Begin VB.Label lblMinimumCycles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Minimum"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   148
         Top             =   1290
         Width           =   840
      End
      Begin VB.Label lblWeightChangeTol 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Weight Change Tolerance:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         TabIndex        =   147
         Top             =   990
         Width           =   2325
      End
      Begin VB.Label lblWeightChangeTolUnits 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "percent"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3270
         TabIndex        =   146
         Top             =   990
         Width           =   670
      End
      Begin VB.Label lblMaximumCycles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   145
         Top             =   1545
         Width           =   840
      End
   End
   Begin VB.Frame frmLeakCheck 
      Caption         =   "Leak Check"
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
      Height          =   1275
      Left            =   120
      TabIndex        =   98
      Top             =   3895
      Width           =   4100
      Begin VB.CheckBox chkLeakCheck 
         Caption         =   "Perform Leak Check First"
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
         Left            =   120
         TabIndex        =   103
         ToolTipText     =   "Check to use Leak Check First"
         Top             =   300
         Width           =   3285
      End
      Begin VB.TextBox txtPauseLeakTime 
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
         Left            =   3230
         MaxLength       =   5
         TabIndex        =   102
         Text            =   "0"
         ToolTipText     =   "0  to 9999 Minutes"
         Top             =   885
         Width           =   715
      End
      Begin VB.CheckBox chkPauseAfterLeak 
         Caption         =   "Pause After LeakCheck     min."
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
         Left            =   120
         TabIndex        =   101
         ToolTipText     =   "Pause After Leak Check"
         Top             =   900
         Width           =   3000
      End
      Begin VB.CheckBox chkLeakPrimary 
         Caption         =   "Primary"
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
         Left            =   120
         MaskColor       =   &H8000000D&
         TabIndex        =   100
         ToolTipText     =   "Leak Check Primary"
         Top             =   600
         Width           =   1605
      End
      Begin VB.CheckBox chkLeakAux 
         Caption         =   "Aux"
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
         Left            =   2400
         MaskColor       =   &H8000000D&
         TabIndex        =   99
         ToolTipText     =   "Leak Check Aux"
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.TextBox txtNotHighlight 
      Height          =   285
      Left            =   11400
      TabIndex        =   33
      Text            =   "NOT Highlight"
      Top             =   13200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frmAuxOutputs 
      BackColor       =   &H80000016&
      Caption         =   "Aux Outputs"
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
      Height          =   2565
      Left            =   4320
      TabIndex        =   52
      Top             =   8700
      Width           =   7095
      Begin VB.Frame frmAuxOutLoad 
         BackColor       =   &H80000016&
         Caption         =   "On during Load"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   1695
         Left            =   120
         TabIndex        =   59
         Top             =   780
         Width           =   3375
         Begin VB.CheckBox chkAuxLoad 
            Caption         =   "Aux Output #4"
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
            TabIndex        =   63
            Top             =   1320
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxLoad 
            Caption         =   "Aux Output #1"
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
            TabIndex        =   62
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxLoad 
            Caption         =   "Aux Output #2"
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
            TabIndex        =   61
            Top             =   680
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxLoad 
            Caption         =   "123456789012345678901234"
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
            TabIndex        =   60
            Top             =   1000
            Width           =   3015
         End
      End
      Begin VB.Frame frmAuxOutPurge 
         BackColor       =   &H80000016&
         Caption         =   "On during Purge"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1695
         Left            =   3600
         TabIndex        =   54
         Top             =   780
         Width           =   3375
         Begin VB.CheckBox chkAuxPurge 
            Caption         =   "Aux Output #3"
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
            TabIndex        =   58
            Top             =   1000
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxPurge 
            Caption         =   "Aux Output #2"
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
            TabIndex        =   57
            Top             =   680
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxPurge 
            Caption         =   "Aux Output #1"
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
            TabIndex        =   56
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox chkAuxPurge 
            Caption         =   "Aux Output #4"
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
            TabIndex        =   55
            Top             =   1320
            Width           =   3015
         End
      End
      Begin VB.CommandButton cmdClose 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6645
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":50890
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Close the Aux Ouputs Configuration Panel"
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Label lblAuxOutMsg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Select Aux Outputs to be Energized"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   390
         Width           =   6015
      End
   End
   Begin VB.Frame frmLoad 
      Caption         =   "Load"
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
      Height          =   6890
      Left            =   8535
      TabIndex        =   12
      Top             =   1570
      Width           =   4140
      Begin VB.Frame frmPostLoad 
         Height          =   945
         Left            =   50
         TabIndex        =   176
         Top             =   5880
         Width           =   4050
         Begin VB.OptionButton optNoPauseAfterLoad 
            Caption         =   "No Pause After Load"
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
            Left            =   120
            TabIndex        =   180
            Top             =   120
            Width           =   3765
         End
         Begin VB.TextBox txtPauseLoadTime 
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
            Left            =   3060
            MaxLength       =   5
            TabIndex        =   179
            Text            =   "0"
            ToolTipText     =   "0  to 9999 Minutes"
            Top             =   360
            Width           =   715
         End
         Begin VB.OptionButton optPauseAfterLoad 
            Caption         =   "Pause After Load      minutes"
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
            Left            =   120
            TabIndex        =   178
            Top             =   390
            Width           =   2925
         End
         Begin VB.OptionButton optPauseAfterLoadForOper 
            Caption         =   "Pause After Load for Operator"
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
            Left            =   120
            TabIndex        =   177
            Top             =   660
            Width           =   3765
         End
      End
      Begin VB.CheckBox chkLoadRatePID 
         Alignment       =   1  'Right Justify
         Caption         =   "Load Rate PID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2290
         TabIndex        =   79
         ToolTipText     =   "Use PID Control of Load Rate"
         Top             =   2615
         Width           =   1605
      End
      Begin VB.Frame frmLiveFuel 
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
         Height          =   1515
         Left            =   45
         TabIndex        =   65
         Top             =   2970
         Visible         =   0   'False
         Width           =   4040
         Begin VB.TextBox txtNitrogenFlow 
            Alignment       =   1  'Right Justify
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
            Left            =   3200
            MaxLength       =   4
            TabIndex        =   71
            Text            =   "0"
            ToolTipText     =   "0 to 5 slpm"
            Top             =   210
            Width           =   715
         End
         Begin VB.CheckBox chkLiveFuel 
            Caption         =   "Use Live Fuel"
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
            Height          =   330
            Left            =   190
            TabIndex        =   70
            ToolTipText     =   "Use Live Fuel"
            Top             =   180
            Width           =   1920
         End
         Begin VB.TextBox txtADF_HeaterSP 
            Alignment       =   1  'Right Justify
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
            Left            =   3200
            MaxLength       =   5
            TabIndex        =   69
            Text            =   "123.4"
            ToolTipText     =   "0 to 45 deg C"
            Top             =   1140
            Width           =   715
         End
         Begin VB.CheckBox chkADF_Heater 
            Caption         =   "Use Fuel Heater"
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
            Height          =   330
            Left            =   190
            TabIndex        =   68
            ToolTipText     =   "Setup Live Fuel"
            Top             =   1110
            Width           =   1920
         End
         Begin VB.TextBox txtLiveFuelChgFreq 
            Alignment       =   1  'Right Justify
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
            Left            =   3200
            TabIndex        =   67
            Text            =   "0"
            ToolTipText     =   "How Often to Change Live Fuel"
            Top             =   510
            Width           =   715
         End
         Begin VB.CheckBox chkLiveFuelChgAuto 
            Caption         =   "Auto Drain/Fill?"
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
            Left            =   190
            TabIndex        =   66
            ToolTipText     =   "Enable Live Fuel Auto Drain/Fill"
            Top             =   885
            Width           =   2875
         End
         Begin VB.Label lblNitrogenFlow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "slpm "
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
            Left            =   2400
            TabIndex        =   75
            Top             =   225
            Width           =   705
         End
         Begin VB.Label lblADF_HeaterSP 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "deg C "
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
            Left            =   2400
            TabIndex        =   74
            Top             =   1155
            Width           =   705
         End
         Begin VB.Label lblLiveFuelChgFreq2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "cycles "
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
            Left            =   2400
            TabIndex        =   73
            Top             =   525
            Width           =   705
         End
         Begin VB.Label lblLiveFuelChgFreq 
            BackStyle       =   0  'Transparent
            Caption         =   "Live Fuel Change Freq:"
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
            Left            =   195
            TabIndex        =   72
            ToolTipText     =   "cycles between fuel changes"
            Top             =   525
            Width           =   1995
         End
      End
      Begin VB.CheckBox chkOrvrMfc 
         Caption         =   "Use ORVR Mfcs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   34
         Top             =   2615
         Width           =   1965
      End
      Begin VB.TextBox txtLoadTime 
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
         Left            =   3180
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "1 to 9999 Minutes"
         Top             =   720
         Width           =   715
      End
      Begin VB.TextBox txtTargetWt 
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
         Left            =   3180
         TabIndex        =   13
         Text            =   "0"
         ToolTipText     =   "0 to 1500 Grams"
         Top             =   1065
         Width           =   715
      End
      Begin VB.TextBox txtWorkCapMult 
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
         Left            =   1860
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   ".1 to 99.9"
         Top             =   2130
         Width           =   615
      End
      Begin VB.TextBox txtLoadBreakthrough 
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
         Left            =   3180
         TabIndex        =   15
         Text            =   "0"
         ToolTipText     =   "0 to 1500 Grams"
         Top             =   1410
         Width           =   715
      End
      Begin VB.TextBox txtEPAFill 
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
         Left            =   3180
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "1 to 36 Hours"
         Top             =   2130
         Width           =   715
      End
      Begin VB.TextBox txtFIDmg 
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
         Left            =   3180
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "0 to 1500 milligrams"
         Top             =   1770
         Visible         =   0   'False
         Width           =   715
      End
      Begin VB.TextBox txtMaxLoadTime 
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
         Left            =   3180
         MaxLength       =   4
         TabIndex        =   27
         Text            =   "0"
         ToolTipText     =   "Load time must exceed load flow devided by canister working capacity Times 60"
         Top             =   5400
         Visible         =   0   'False
         Width           =   715
      End
      Begin VB.TextBox txtLoadRate 
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
         Left            =   3180
         TabIndex        =   26
         Text            =   "0"
         ToolTipText     =   "Enter a value in grams per hour."
         Top             =   4680
         Width           =   715
      End
      Begin VB.TextBox txtButnPercent 
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
         Left            =   3180
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "0.0"
         ToolTipText     =   "1 to 100 Percent"
         Top             =   5040
         Width           =   715
      End
      Begin VB.CheckBox optFIDBreakthrough 
         Caption         =   "FID Breakthrough             mg"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   24
         Top             =   1755
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.CheckBox optNoLoad 
         Caption         =   "No Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   3645
      End
      Begin VB.CheckBox optLoadTime 
         Caption         =   "Load by Time             minutes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   22
         Top             =   705
         Width           =   3645
      End
      Begin VB.CheckBox optWCM 
         Caption         =   "Work Cap Mul               hours"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   21
         Top             =   2100
         Width           =   3645
      End
      Begin VB.CheckBox optLoadweight 
         Caption         =   "Load by Weight            grams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   20
         Top             =   1050
         Width           =   3645
      End
      Begin VB.CheckBox optLoadBreakthrough 
         Caption         =   "Load by Breakthrough   grams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   19
         Top             =   1395
         Width           =   3645
      End
      Begin VB.Label lblMaxLoadTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Load Time        minutes"
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
         Left            =   240
         TabIndex        =   32
         Top             =   5415
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Target Load Rate:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   31
         Top             =   4710
         Width           =   1815
      End
      Begin VB.Label lblButnPercent 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent Butane:"
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
         Left            =   240
         TabIndex        =   30
         Top             =   5055
         Width           =   1935
      End
      Begin VB.Label lblButnPercentUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "percent "
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
         Left            =   2190
         TabIndex        =   29
         Top             =   5055
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " grams/hr "
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
         Left            =   2070
         TabIndex        =   28
         Top             =   4688
         Width           =   1080
      End
   End
   Begin VB.Frame frmPurge 
      Caption         =   "Purge"
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
      Height          =   6890
      Left            =   4320
      TabIndex        =   7
      Top             =   1570
      Width           =   4140
      Begin VB.Frame frmPostPurge 
         Height          =   945
         Left            =   50
         TabIndex        =   172
         Top             =   5880
         Width           =   4050
         Begin VB.OptionButton optNoPauseAfterPurge 
            Caption         =   "No Pause After Purge"
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
            Left            =   120
            TabIndex        =   181
            Top             =   120
            Width           =   3765
         End
         Begin VB.OptionButton optPauseAfterPurgeForOper 
            Caption         =   "Pause After Purge for Operator"
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
            Left            =   120
            TabIndex        =   175
            Top             =   660
            Width           =   3765
         End
         Begin VB.OptionButton optPauseAfterPurge 
            Caption         =   "Pause After Purge      minutes"
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
            Left            =   120
            TabIndex        =   174
            Top             =   390
            Width           =   2925
         End
         Begin VB.TextBox txtPausePurgeTime 
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
            Left            =   3060
            MaxLength       =   5
            TabIndex        =   173
            Text            =   "0"
            ToolTipText     =   "0  to 9999 Minutes"
            Top             =   360
            Width           =   715
         End
      End
      Begin VB.CheckBox chkUsePurgeOven 
         Caption         =   "Use Oven"
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
         Left            =   240
         TabIndex        =   170
         ToolTipText     =   "Canister in Oven during Purge"
         Top             =   5100
         Width           =   1920
      End
      Begin VB.TextBox txtPurgeOvenSP 
         Alignment       =   1  'Right Justify
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
         Left            =   3180
         MaxLength       =   5
         TabIndex        =   169
         Text            =   "123.4"
         ToolTipText     =   "0 to 60 deg C"
         Top             =   5070
         Width           =   715
      End
      Begin VB.CheckBox chkPurgeCansInSeries 
         Alignment       =   1  'Right Justify
         Caption         =   "Purge Canisters in Series"
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
         Left            =   210
         TabIndex        =   165
         ToolTipText     =   "Purge the Primary and Aux Canisters in Series "
         Top             =   5685
         Width           =   3630
      End
      Begin VB.CommandButton cmdPurgeProfile 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":50D82
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Open the Purge Profile Configuration screen"
         Top             =   1110
         UseMaskColor    =   -1  'True
         Width           =   715
      End
      Begin VB.Frame frmPurgeTargetMode 
         Height          =   3135
         Left            =   40
         TabIndex        =   82
         Top             =   1470
         Width           =   4040
         Begin VB.TextBox txtPurgeVolume 
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
            Left            =   3135
            MaxLength       =   4
            TabIndex        =   96
            Text            =   "1"
            ToolTipText     =   "0.01 to 9999"
            Top             =   495
            Width           =   715
         End
         Begin VB.TextBox txtPurgeLiters 
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
            Left            =   3135
            MaxLength       =   4
            TabIndex        =   182
            Text            =   "1"
            ToolTipText     =   "0.01 to 9999"
            Top             =   235
            Width           =   715
         End
         Begin VB.CheckBox optPurgeLiters 
            Caption         =   "Purge by Liters               liters"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   200
            TabIndex        =   183
            Top             =   220
            Width           =   3645
         End
         Begin VB.CommandButton cmdPurgeWizard 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRecipe.frx":51674
            Style           =   1  'Graphical
            TabIndex        =   168
            ToolTipText     =   "Open Purge Calculations Wizard"
            Top             =   2610
            UseMaskColor    =   -1  'True
            Width           =   715
         End
         Begin VB.TextBox txtPurgeWC 
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
            Left            =   3135
            TabIndex        =   94
            Text            =   "0"
            ToolTipText     =   "Target Weight Loss in % of Canister Working Capacity"
            Top             =   765
            Width           =   715
         End
         Begin VB.CheckBox optPurgeVolume 
            Caption         =   "Purge by Volumes      volumes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   200
            TabIndex        =   97
            Top             =   480
            Width           =   3645
         End
         Begin VB.CheckBox optPurgeUndo 
            Caption         =   "Purge to Undo Load"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   200
            TabIndex        =   95
            Top             =   1290
            Width           =   3645
         End
         Begin VB.TextBox txtPurgeTarget 
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
            Left            =   3135
            TabIndex        =   91
            Text            =   "0"
            ToolTipText     =   "Target Weight in grams"
            Top             =   1035
            Width           =   715
         End
         Begin VB.CheckBox optPurgeTarget 
            Caption         =   "Purge to Target            grams"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   200
            TabIndex        =   93
            Top             =   1020
            Width           =   3645
         End
         Begin VB.CheckBox optPurgeWC 
            Caption         =   "Purge by Work Cap   % of WC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   200
            TabIndex        =   92
            Top             =   750
            Width           =   3645
         End
         Begin VB.OptionButton optTargetContinuous 
            Caption         =   "Continuous Purge"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   455
            TabIndex        =   87
            ToolTipText     =   "Purge to Target at a Constant Flow Rate"
            Top             =   1950
            Width           =   2715
         End
         Begin VB.OptionButton optTargetPurgePauseRepeat 
            Caption         =   "Purge-Pause Cycles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   455
            TabIndex        =   86
            ToolTipText     =   "Cyclically Purge-Pause to Target"
            Top             =   2190
            Width           =   3375
         End
         Begin VB.TextBox txtTargetTimeout 
            Alignment       =   1  'Right Justify
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
            Left            =   3135
            TabIndex        =   85
            Text            =   "0"
            ToolTipText     =   "Maximum Canister Volumes per Purge"
            Top             =   1695
            Width           =   715
         End
         Begin VB.TextBox txtTargetPurge 
            Alignment       =   1  'Right Justify
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
            Left            =   3135
            TabIndex        =   84
            Text            =   "0"
            ToolTipText     =   "minutes to purge each PurgePause cycle"
            Top             =   2400
            Width           =   715
         End
         Begin VB.TextBox txtTargetPause 
            Alignment       =   1  'Right Justify
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
            Left            =   3135
            TabIndex        =   83
            Text            =   "0"
            ToolTipText     =   "minutes to pause each PurgePause cycle"
            Top             =   2685
            Width           =   715
         End
         Begin VB.Label lblTargetTimeoutUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "max volumes "
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
            Left            =   1950
            TabIndex        =   90
            Top             =   1710
            Width           =   1905
         End
         Begin VB.Label lblTargetPurgeUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "purge (min) "
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
            Left            =   1995
            TabIndex        =   89
            Top             =   2415
            Width           =   1845
         End
         Begin VB.Label lblTargetPauseUnits 
            BackStyle       =   0  'Transparent
            Caption         =   "pause (min) "
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
            Left            =   1995
            TabIndex        =   88
            Top             =   2700
            Width           =   1845
         End
      End
      Begin VB.TextBox txtPurgeAuxOnly 
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
         Left            =   3180
         TabIndex        =   37
         Text            =   "0"
         ToolTipText     =   "1 to 9999 Minutes"
         Top             =   855
         Width           =   715
      End
      Begin VB.TextBox txtPurgeProfile 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2380
         TabIndex        =   78
         Text            =   "0"
         ToolTipText     =   "Profile Number 1-99; 0 = none/na"
         Top             =   1200
         Width           =   715
      End
      Begin VB.CheckBox optPurgeProfile 
         Caption         =   "Purge by Profile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   77
         Top             =   1110
         Width           =   3645
      End
      Begin VB.CheckBox optPurgeAuxOnly 
         Caption         =   "Purge Aux Only          minutes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   3645
      End
      Begin VB.TextBox txtPurgeTime 
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
         Left            =   3180
         TabIndex        =   35
         Text            =   "0"
         ToolTipText     =   "1 to 9999 Minutes"
         Top             =   585
         Width           =   715
      End
      Begin VB.CheckBox optPurgeTime 
         Caption         =   "Purge by Time           minutes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   36
         Top             =   570
         Width           =   3645
      End
      Begin VB.TextBox txtPurgeFlow 
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
         Left            =   3180
         MaxLength       =   5
         TabIndex        =   10
         Text            =   "1"
         ToolTipText     =   "Enter a value in slpm of less than 95% full range."
         Top             =   4680
         Width           =   715
      End
      Begin VB.CheckBox chkPurgeAuxCan 
         Alignment       =   1  'Right Justify
         Caption         =   "Purge Auxiliary Canister"
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
         Left            =   210
         TabIndex        =   9
         ToolTipText     =   "Use Auxiliary Canister"
         Top             =   5415
         Width           =   3630
      End
      Begin VB.CheckBox optNoPurge 
         Caption         =   "No Purge"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   3645
      End
      Begin VB.Label lblPurgeOvenUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "deg C "
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
         Left            =   2445
         TabIndex        =   171
         Top             =   5100
         Width           =   705
      End
      Begin VB.Label lblPurgeFlow 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Flow Rate:                slpm"
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
         Left            =   240
         TabIndex        =   11
         Top             =   4710
         Width           =   2925
      End
   End
   Begin VB.Frame frmCycle 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   3890
      Left            =   120
      TabIndex        =   6
      Top             =   7375
      Width           =   4100
      Begin VB.Frame frmLineVolume 
         Caption         =   "Line Volume"
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
         Height          =   1440
         Left            =   50
         TabIndex        =   112
         Top             =   2400
         Visible         =   0   'False
         Width           =   4000
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
            TabIndex        =   121
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
            TabIndex        =   120
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
            TabIndex        =   119
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
            TabIndex        =   118
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
            TabIndex        =   117
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
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
            Text            =   "0.0"
            ToolTipText     =   "Calculated Volume in liters"
            Top             =   500
            Width           =   615
         End
         Begin VB.Label Label30 
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
            TabIndex        =   136
            Top             =   525
            Width           =   500
         End
         Begin VB.Label Label29 
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
            TabIndex        =   135
            Top             =   795
            Width           =   500
         End
         Begin VB.Label Label22 
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
            TabIndex        =   134
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
            TabIndex        =   133
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
            TabIndex        =   132
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
            TabIndex        =   131
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
            TabIndex        =   130
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
            TabIndex        =   129
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
            TabIndex        =   128
            Top             =   525
            Width           =   285
         End
         Begin VB.Label Label21 
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
            TabIndex        =   127
            Top             =   280
            Width           =   1100
         End
         Begin VB.Label Label20 
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
            TabIndex        =   126
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label31 
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
            TabIndex        =   125
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
            TabIndex        =   124
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
            TabIndex        =   123
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
            TabIndex        =   122
            Top             =   525
            Width           =   455
         End
      End
      Begin VB.Frame frmResources 
         Caption         =   "Resources"
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
         Height          =   2205
         Left            =   50
         TabIndex        =   104
         Top             =   120
         Width           =   4000
         Begin VB.CheckBox chkCommonTC 
            Caption         =   "Use Common Thermocouples"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   111
            ToolTipText     =   "Check to end with Purge on multiple Purge/Load or Load Purge cycles"
            Top             =   916
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.CheckBox chkUseAuxScale 
            Caption         =   "Check to use Auxiliary Scale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   110
            ToolTipText     =   "Use Auxiliary Scale"
            Top             =   593
            Width           =   2865
         End
         Begin VB.CheckBox chkPrimaryScale 
            Caption         =   "Check to use Primary Scale"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   109
            ToolTipText     =   "Use Primary Scale"
            Top             =   270
            Width           =   2775
         End
         Begin VB.TextBox txtPrimaryScaleNo 
            Alignment       =   1  'Right Justify
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
            Left            =   3180
            MaxLength       =   3
            TabIndex        =   108
            Text            =   "0"
            ToolTipText     =   "0 to Max Scales"
            Top             =   270
            Width           =   715
         End
         Begin VB.TextBox txtAuxScaleNo 
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
            Left            =   3180
            MaxLength       =   3
            TabIndex        =   107
            Text            =   "0"
            ToolTipText     =   "0 to Max Scales"
            Top             =   616
            Width           =   715
         End
         Begin VB.CheckBox chkAuxOutputs 
            Caption         =   "Use Auxiliary Outputs"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   120
            TabIndex        =   106
            ToolTipText     =   "Check to use auxiliary 12vdc or Dry Contact outputs"
            Top             =   1239
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.CommandButton cmdCfgAuxOutputs 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3180
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmRecipe.frx":519B6
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Open the Aux Ouputs Configuration Panel"
            Top             =   1239
            UseMaskColor    =   -1  'True
            Width           =   735
         End
      End
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   10560
      Top             =   13440
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
      Height          =   2730
      Left            =   4320
      TabIndex        =   4
      Top             =   8535
      Width           =   8355
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
         Height          =   2385
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   8130
      End
   End
   Begin VB.Frame frmHighlight 
      BackColor       =   &H8000000D&
      Caption         =   "highlight"
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   11160
      TabIndex        =   2
      Top             =   13680
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame frmNotHighlight 
      Caption         =   "NOT highlight"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11040
      TabIndex        =   3
      Top             =   13440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Threed.SSPanel pnlRecipe 
      Height          =   510
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   12525
      _Version        =   65536
      _ExtentX        =   22093
      _ExtentY        =   900
      _StockProps     =   15
      Caption         =   "  Name: "
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
      Begin VB.TextBox txtRecipeName 
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
         MaxLength       =   100
         TabIndex        =   1
         ToolTipText     =   "Alphanumeric Description "
         Top             =   120
         Width           =   10515
      End
   End
   Begin VB.PictureBox pbControlBtns 
      Align           =   1  'Align Top
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12840
      TabIndex        =   39
      Top             =   0
      Width           =   12840
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         DisabledPicture =   "frmRecipe.frx":51EA8
         DownPicture     =   "frmRecipe.frx":525AA
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
         Left            =   4560
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":52CAC
         Style           =   1  'Graphical
         TabIndex        =   167
         ToolTipText     =   "Reload Station Recipe Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPrintAll 
         Caption         =   "Print All"
         DisabledPicture =   "frmRecipe.frx":533AE
         DownPicture     =   "frmRecipe.frx":53FF0
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
         Picture         =   "frmRecipe.frx":54C32
         Style           =   1  'Graphical
         TabIndex        =   80
         ToolTipText     =   "Print All Master Recipes"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         DisabledPicture =   "frmRecipe.frx":55874
         DownPicture     =   "frmRecipe.frx":55F76
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
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":56678
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Open Master Recipe List"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         DisabledPicture =   "frmRecipe.frx":56D7A
         DownPicture     =   "frmRecipe.frx":5747C
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
         Left            =   5400
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":57B7E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Copy Recipe Values to the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Paste"
         DisabledPicture =   "frmRecipe.frx":58280
         DownPicture     =   "frmRecipe.frx":58982
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
         Left            =   6240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":59084
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Paste Recipe Values from the clipboard"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         DisabledPicture =   "frmRecipe.frx":59786
         DownPicture     =   "frmRecipe.frx":59E88
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
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":5A58A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Save Recipe"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "Restore"
         DisabledPicture =   "frmRecipe.frx":5AC8C
         DownPicture     =   "frmRecipe.frx":5B38E
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
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":5BA90
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Reload Station Recipe Values"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         DisabledPicture =   "frmRecipe.frx":5C192
         DownPicture     =   "frmRecipe.frx":5CDD4
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
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":5DA16
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Print current Recipe"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdreturn 
         Caption         =   "Close"
         DisabledPicture =   "frmRecipe.frx":5E658
         DownPicture     =   "frmRecipe.frx":5ED5A
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
         Left            =   11820
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":5F45C
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Quit"
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPgDn 
         Caption         =   "Pg Prev"
         DisabledPicture =   "frmRecipe.frx":5FB5E
         DownPicture     =   "frmRecipe.frx":60260
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
         Left            =   7320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":60962
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   "Prev"
         DisabledPicture =   "frmRecipe.frx":61064
         DownPicture     =   "frmRecipe.frx":61766
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
         Left            =   8160
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":61E68
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "Next"
         DisabledPicture =   "frmRecipe.frx":6256A
         DownPicture     =   "frmRecipe.frx":62C6C
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
         Left            =   10005
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":6336E
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdPgUp 
         Caption         =   "Pg Next"
         DisabledPicture =   "frmRecipe.frx":63A70
         DownPicture     =   "frmRecipe.frx":64172
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
         Left            =   10845
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmRecipe.frx":64874
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin Threed.SSPanel pnlDispRcpNum 
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
         Left            =   9000
         TabIndex        =   51
         ToolTipText     =   "Click for list of Defined Recipes"
         Top             =   0
         Width           =   1005
         _Version        =   65536
         _ExtentX        =   1773
         _ExtentY        =   1482
         _StockProps     =   15
         Caption         =   "49"
         ForeColor       =   -2147483646
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
         ForeColor       =   &H80000002&
         Height          =   345
         Left            =   7170
         TabIndex        =   166
         Top             =   240
         Visible         =   0   'False
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmRecipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' error module 90 ''''''''''''' Form RECIPE.frm ''''''''''''''''''''''''
Option Explicit
Private RecipeMode As Integer            ' 0=master; 1=station
Private DispRcp As Integer               ' Current Master Recipe index
Private ScreenBkgdColor As Long
Private ScreenDescription As String
Private ScreenDispFlag As Boolean
Private StnShftDescription As String
Private Chgs As Boolean
Private DspRecipe As Recipe
Private MemRecipe As Recipe
Private inct As Integer
Private dbDbase As Database
Private rsRecord  As Recordset
Private Criteria As String
Private sdate As Date
Private sSec As Double
Private tmpval1, tmpval2, tmpval3 As Single
Private tmpStr As String
Private Const displayNone = 0
Private Const displayNoMaxVol = 1
Private Const displayAll = 2

Public Sub ChgRecipeMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    RecipeMode = IIf((NewMode = MASTERMODE Or NewMode = STATIONMODE), NewMode, MASTERMODE)
    Select Case RecipeMode
        Case MASTERMODE
            ' station/shift description
            lblStnDesc.Visible = False
            ' screen description
            ScreenDescription = "Master Recipe Parameters"
            ' screen background color
            ScreenBkgdColor = MasterMode_BackColor
            ' show MasterOnly items
            ScreenDispFlag = True
            ' update screen
            UpdateRecipeScreen
        Case STATIONMODE
            ' station/shift description
            StnShftDescription = "Station #" & Format(DispStn, "#0")
            If NR_SHIFT > 1 Then StnShftDescription = StnShftDescription & "  Shift #" & Format(DispShift, "#0")
            StnShftDescription = StnShftDescription & "  Recipe Parameters"
            lblStnDesc.Visible = True
'            lblStnDesc.Left = cmdPgDn.Left
            lblStnDesc.ForeColor = TitlesData_Forecolor
            lblStnDesc.Caption = StnShftDescription
            ' screen description
            ScreenDescription = StnShftDescription
            ' screen background color
            ScreenBkgdColor = StationMode_BackColor
            ' hide MasterOnly items
            ScreenDispFlag = False
            ' update screen
            UpdateRecipeScreen
    End Select
    ' screen description
    frmRecipe.Caption = ScreenDescription
    ' set screen background colors
    frmRecipe.BackColor = ScreenBkgdColor
    pbControlBtns.BackColor = ScreenBkgdColor
    pnlDispRcpNum.BackColor = ScreenBkgdColor
    ' show Recipe # & Arrows ??
    cmdDown.Visible = ScreenDispFlag
    cmdUp.Visible = ScreenDispFlag
    cmdPgDn.Visible = ScreenDispFlag
    cmdPgUp.Visible = ScreenDispFlag
    pnlDispRcpNum.Visible = ScreenDispFlag
End Sub

Public Sub LoadNewRcp(ByVal NewRcp As Integer)
    DispRcp = NewRcp
    frmRecipe.tmrUpdate.Enabled = True
    RecipeDisplay_ByNum
End Sub

Public Sub SetPurgeProfile(ByVal NewProf As Integer)
    txtPurgeProfile.text = "#" & Format(NewProf, "##0")
End Sub

Public Function OkToRunRecipeInStation() As Boolean
    OkToRunRecipeInStation = ValidRecipe
End Function

Public Sub ExportRecipe()
    ScreenToDspRcp
    ExportedRecipe = DspRecipe
End Sub

Public Sub UpdatePurgeFlowTime(ByVal newTime As Single, ByVal newFlow As Single)
    txtPurgeTime.text = Format(newTime, "###,##0.##")
    txtPurgeFlow.text = Format(newFlow, "#,##0.##")
End Sub

Private Sub DspRcpToScreen()
' Copies DspRecipe to Screen
Dim optDisplay As Integer
Dim Idx As Integer
Dim sTxt As String
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 9
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    ' show Recipe # & name
    pnlDispRcpNum.Caption = Format(DspRecipe.Number, "#00")
    txtRecipeName.text = DspRecipe.Name

    ' cycle options
    tabsCycletype.Tabs(DspRecipe.CycleType).Selected = True
    ' purge options
    Select Case DspRecipe.Purge_Method
        Case NOPURGE
            optNoPurge.Value = cYES
'            optDisplay = displayNone
        Case PURGEBYTIME
            optPurgeTime.Value = cYES
'            optDisplay = displayNone
        Case PURGEBYLITERS
            optPurgeLiters.Value = cYES
'            optDisplay = displayNoMaxVol
        Case PURGEBYVOLUME
            optPurgeVolume.Value = cYES
'            optDisplay = displayNoMaxVol
        Case PURGEAUXONLY
            optPurgeAuxOnly.Value = cYES
'            optDisplay = displayNone
        Case PURGEBYPROFILE
            optPurgeProfile.Value = cYES
'            optDisplay = displayNone
        Case PURGEBYWC
            optPurgeWC.Value = cYES
'            optDisplay = displayAll
       Case PURGETOTARGET
            optPurgeTarget.Value = cYES
'            optDisplay = displayAll
        Case PURGETOUNDOLOAD
            optPurgeUndo.Value = cYES
'            optDisplay = displayAll
        Case Else
            ' no purge
            optNoPurge.Value = cYES
'            optDisplay = displayNone
    End Select
    ' purge-to-target options
    Select Case DspRecipe.Purge_TargetMode
        Case NOTARGET
            optTargetContinuous.Value = True
        Case TARGETCONTINUOUS
            optTargetContinuous.Value = True
        Case TARGETPURGEPAUSE
            optTargetPurgePauseRepeat.Value = True
        Case Else
            ' no purge-to-target
            optTargetContinuous.Value = True
    End Select
    txtTargetTimeout.text = Format(DspRecipe.Purge_MaxVolumes, "###0")
    txtTargetPurge.text = Format(DspRecipe.Purge_TargetPurge, "##0.0##")
    txtTargetPause.text = Format(DspRecipe.Purge_TargetPause, "##0.0##")
    
    txtPurgeAuxOnly = DspRecipe.Purge_AuxTime
    txtPurgeTime.text = Format(DspRecipe.Purge_Time, "####0")
    txtPurgeFlow = DspRecipe.Purge_Flow
    txtPurgeLiters = DspRecipe.Purge_Liters
    txtPurgeVolume = DspRecipe.Purge_Can_Vol
    txtPurgeProfile.text = "#" & Format(DspRecipe.Purge_ProfileNumber, "##0")
    txtPurgeTarget.text = Format(DspRecipe.Purge_TargetWeight, "####0.0##")
    txtPurgeWC.text = Format(DspRecipe.Purge_TargetWC, "####0.0##")
    
    optNoLoad.Value = IIf(DspRecipe.Load_Method = NOLOAD, cYES, cNO)
    optLoadTime.Value = IIf(DspRecipe.Load_Method = LOADBYTIME, cYES, cNO)
    optLoadweight.Value = IIf(DspRecipe.Load_Method = LOADBYWEIGHT, cYES, cNO)
    optLoadBreakthrough.Value = IIf(DspRecipe.Load_Method = LOADBYBREAKTHRU, cYES, cNO)
    optFIDBreakthrough.Value = IIf(DspRecipe.Load_Method = LOADBYFID, cYES, cNO)
    optWcm.Value = IIf(DspRecipe.Load_Method = LOADBYWC, cYES, cNO)
    
    txtEPAFill = DspRecipe.EPAFill
    txtLoadRate = DspRecipe.Load_Rate
    txtButnPercent = DspRecipe.Mix_Percent
    txtWorkCapMult = Format(DspRecipe.WC_MultSave, "#0.0#")
    txtLoadTime = DspRecipe.Load_Time
    txtTargetWt = DspRecipe.Load_Wt
    txtLoadBreakthrough = DspRecipe.LoadBreakthrough

'    chkUseAnalyzer.Value = IIf(DspRecipe.UseAnalyzer, cYES, cNO)
    txtFIDmg = Format(DspRecipe.FIDmg, "###0")
    txtPauseLeakTime = DspRecipe.PauseLeakTime
    txtPauseLoadTime = DspRecipe.PauseLoadTime
    txtPausePurgeTime = DspRecipe.PausePurgeTime
    optPauseAfterPurge.Value = IIf(DspRecipe.PauseAfterPurge, True, False)
    optPauseAfterPurgeForOper.Value = IIf(DspRecipe.PauseAfterPurgeForOper, True, False)
    optNoPauseAfterPurge.Value = IIf((DspRecipe.PauseAfterPurge Or DspRecipe.PauseAfterPurgeForOper), False, True)

    chkPurgeAuxCan.Value = IIf(DspRecipe.PurgeAuxCan, cYES, cNO)
    chkPurgeCansInSeries.Value = IIf(DspRecipe.PurgeCansInSeries, cYES, cNO)
    optPauseAfterLoad.Value = IIf(DspRecipe.PauseAfterLoad, True, False)
    optPauseAfterLoadForOper.Value = IIf(DspRecipe.PauseAfterLoadForOper, True, False)
    optNoPauseAfterLoad.Value = IIf((DspRecipe.PauseAfterLoad Or DspRecipe.PauseAfterLoadForOper), False, True)
    
    chkUsePurgeOven.Value = IIf(DspRecipe.PurgeOven, cYES, cNO)
    txtPurgeOvenSP.text = Format(DspRecipe.PurgeOvenSP, "###0.0##")
    txtNitrogenFlow.text = Format(DspRecipe.NitrogenFlowSave, "###0.0##")
'    frmAnalyzer.txtTargetConcentration = DspRecipe.TargetConcentration
'    frmAnalyzer.txtDwellTime = DspRecipe.DwellTime
    txtRecipeName.text = DspRecipe.Name
    txtMaxLoadTime.text = DspRecipe.MaxLoadTime
    chkOrvrMfc.Value = IIf(DspRecipe.UseHiRangeMFC, cYES, cNO)

    txtAuxScaleNo.text = Format(DspRecipe.AuxScaleNo, "#0")
    txtPrimaryScaleNo.text = Format(DspRecipe.PriScaleNo, "#0")
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        chkPrimaryScale.Value = cYES
        chkUseAuxScale.Value = cYES
    Else
        chkPrimaryScale.Value = IIf(DspRecipe.UsePriScale, cYES, cNO)
        chkUseAuxScale.Value = IIf(DspRecipe.UseAuxScale, cYES, cNO)
    End If

    ' line volume options
    txtIDLoad = CStr(DspRecipe.IDLoad)
    txtLoadL = CStr(DspRecipe.LoadL)
    txtLoadV = CStr(DspRecipe.LoadV)
    txtIDPurge = CStr(DspRecipe.IDPurge)
    txtPurgeL = CStr(DspRecipe.PurgeL)
    txtPurgeV = CStr(DspRecipe.PurgeV)
    txtIDVent = CStr(DspRecipe.IDVent)
    txtVentL = CStr(DspRecipe.VentL)
    txtVentV = CStr(DspRecipe.VentV)

    ' leakcheck options
    chkLeakCheck = IIf(DspRecipe.LeakCheck, cYES, cNO)
    If USINGAUXLEAKCHECK Then
        chkLeakPrimary.Value = IIf(DspRecipe.LeakPrimary, cYES, cNO)
        chkLeakAux.Value = IIf(DspRecipe.LeakAux, cYES, cNO)
        chkLeakPrimary.Visible = True
        chkLeakAux.Visible = True
    Else
        chkLeakPrimary.Value = cYES
        chkLeakAux.Value = cNO
        chkLeakPrimary.Visible = False
        chkLeakAux.Visible = False
    End If
    chkPauseAfterLeak.Value = IIf(DspRecipe.PauseAfterLeak, cYES, cNO)
    
    ' start options
    Select Case DspRecipe.StartMethod
        Case STARTNOW
            optStartNow.Value = cON
        Case STARTDELAYED
            optStartAfter.Value = cON
        Case STARTATDATE
            optStartAt.Value = cON
        Case Else
            ' start with no delay
            optStartNow.Value = cON
    End Select
    txtStartAfterMin.text = Format(DspRecipe.StartDelay, "##0")
    txtStartAtDate.text = Format(DspRecipe.StartDate, "yyyy-MMM-dd hh:mm")
    
    ' end options
    Select Case DspRecipe.EndMethod
        Case ENDCYCLES
            ' end after n cycles
            optEndCycles.Value = cON
        Case ENDWEIGHTCHG
            ' end after stable weight change
            optEndWeightChange.Value = cON
        Case Else
            ' end after n cycles
            optEndCycles.Value = cON
    End Select
    optUpdateCanWc.Value = IIf(DspRecipe.UpdateCanWc, cON, cOFF)
    txtPFCycle.text = Format(DspRecipe.CyclesSave, "##0")
    txtWeightChangeTol.text = Format(DspRecipe.EndWeightTolerance, "##0.0#")
    txtConsecutiveCycles.text = Format(DspRecipe.EndConsecutiveCycles, "##0")
    txtMaximumCycles.text = Format(DspRecipe.EndMaximumCycles, "##0")
    txtMinimumCycles.text = Format(DspRecipe.EndMinimumCycles, "##0")

    If (IsEmpty(txtFIDmg.text)) Then txtFIDmg = "0"
    If (IsNull(txtFIDmg.text)) Then txtFIDmg = "0"
    If (Not IsNumeric(txtFIDmg.text)) Then txtFIDmg = "0"
    
    ' live fuel options
    txtNitrogenFlow.text = Format(DspRecipe.NitrogenFlowSave, "###0.0##")
    txtNitrogenFlow.ToolTipText = Format(DspRecipe.NitrogenFlowSave, "###0.0##")
    txtLiveFuelChgFreq = DspRecipe.LiveFuelChgFreq
    chkLiveFuelChgAuto.Value = IIf(DspRecipe.LiveFuelChgAuto, cYES, cNO)
    chkADF_Heater.Value = IIf(DspRecipe.ADF_Heater, cYES, cNO)
    txtADF_HeaterSP.text = Format(DspRecipe.ADF_HeaterSP, "##0.0")
    chkLoadRatePID = IIf(DspRecipe.UseLoadRatePID, cYES, cNO)
    chkLoadRatePID.Visible = IIf(DspRecipe.LiveFuel, True, False)
    
    ' aux outputs
    chkAuxOutputs.Value = IIf(DspRecipe.AuxOutputs, cON, cOFF)
    For Idx = 1 To 4
        chkAuxLoad(Idx).Value = IIf(DspRecipe.AuxOutputs_Load(Idx), cON, cOFF)
        chkAuxPurge(Idx).Value = IIf(DspRecipe.AuxOutputs_Purge(Idx), cON, cOFF)
    Next Idx
    ' aux outputs selection button
    cmdCfgAuxOutputs.Visible = IIf(((USING_AUX_OUTPUTS) And (chkAuxOutputs.Value = cON)), True, False)
    
    Select Case RecipeMode
        Case MASTERMODE
            If DspRecipe.LiveFuel Then
                chkLiveFuel.Value = cYES
                chkOrvrMfc.Visible = False
                chkOrvrMfc.Value = cNO
                txtButnPercent.Visible = False
                lblButnPercent.Visible = False
                lblButnPercentUnits.Visible = False
                txtNitrogenFlow.ToolTipText = "0.1 to 50 slpm"
            Else
                chkLiveFuel.Value = cNO
                chkOrvrMfc.Visible = IIf(systemhasORVR2, True, False)
                txtButnPercent.Visible = True
                lblButnPercent.Visible = True
                lblButnPercentUnits.Visible = True
                txtNitrogenFlow.ToolTipText = "0.1 to 50 slpm"
            End If
        Case STATIONMODE
            Select Case STN_INFO(DispStn).Type
                Case STN_REGULAR_TYPE
'                    DspRecipe.LiveFuel = False
                    frmLiveFuel.Visible = False
                    chkOrvrMfc.Visible = False
                    txtButnPercent.Visible = True
                    lblButnPercent.Visible = True
                    lblButnPercentUnits.Visible = True
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asNitrogenFlowSP).EuMax, "##0.0") & " slpm"
                Case STN_ORVR_TYPE
'                    DspRecipe.LiveFuel = False
                    frmLiveFuel.Visible = False
                    chkOrvrMfc.Visible = False
                    txtButnPercent.Visible = True
                    lblButnPercent.Visible = True
                    lblButnPercentUnits.Visible = True
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asNitrogenFlowSP).EuMax, "##0.0") & " slpm"
                Case STN_ORVR2_TYPE
'                    DspRecipe.LiveFuel = False
                    frmLiveFuel.Visible = False
                    chkOrvrMfc.Visible = True
                    txtButnPercent.Visible = True
                    lblButnPercent.Visible = True
                    lblButnPercentUnits.Visible = True
                    If DspRecipe.UseHiRangeMFC Then
                        ' hi-range MFC
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax, "##0.0") & " slpm"
                    Else
                        ' lo-range MFC
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asNitrogenFlowSP).EuMax, "##0.0") & " slpm"
                    End If
                Case STN_LIVEFUEL_TYPE
'                    DspRecipe.LiveFuel = True
                    frmLiveFuel.Visible = True
                    chkLoadRatePID.Visible = True
                    If (STN_INFO(DispStn).ADF_TANKTYPE > 10 And STN_INFO(DispStn).ADF_TANKTYPE <= 20) Then
                        chkADF_Heater.Enabled = True
                        chkADF_Heater.Visible = True
                        chkADF_Heater.Caption = "Use Fuel Heater"
                        lblADF_HeaterSP.Visible = True
                        If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                        If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                        txtADF_HeaterSP.Enabled = True
                        txtADF_HeaterSP.Visible = True
                        If USINGC Then txtADF_HeaterSP.ToolTipText = "15 to 50 deg C"
                        If USINGF Then txtADF_HeaterSP.ToolTipText = "60 to 120 deg F"
                    ElseIf (STN_INFO(DispStn).ADF_TANKTYPE = 90) Then
                        chkADF_Heater.Enabled = True
                        chkADF_Heater.Visible = True
                        chkADF_Heater.Caption = "Use WaterBath"
                        lblADF_HeaterSP.Visible = True
                        If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                        If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                        txtADF_HeaterSP.Enabled = True
                        txtADF_HeaterSP.Visible = True
                        Select Case Idx
                            Case wbDirect
                                sTxt = "WaterBath SetPoint from "
                            Case wbFuelTemp
                                sTxt = "LiveFuel Fuel SetPoint from "
                            Case wbVaporTemp
                                sTxt = "LiveFuel Vapor SetPoint from "
                        End Select
                        If USINGC Then txtADF_HeaterSP.ToolTipText = sTxt & Format(WB_AIO.EuMin, "###0.0##") & " to " & Format(WB_AIO.EuMax, "###0.0##") & " deg C"
                        If USINGF Then txtADF_HeaterSP.ToolTipText = sTxt & Format(DegCtoF(WB_AIO.EuMin), "###0.0##") & " to " & Format(DegCtoF(WB_AIO.EuMax), "###0.0##") & " deg F"
                    Else
                        chkADF_Heater.Visible = False
                        lblADF_HeaterSP.Visible = False
                        txtADF_HeaterSP.Visible = False
                    End If
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax, "##0.#") & " slpm"
                    chkOrvrMfc.Visible = False
                    txtButnPercent.Visible = False
                    lblButnPercent.Visible = False
                    lblButnPercentUnits.Visible = False
                Case STN_LIVEREG_TYPE
                    frmLiveFuel.Visible = True
                    chkOrvrMfc.Visible = False
                    If DspRecipe.LiveFuel Then
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax, "##0.#") & " slpm"
                        If (STN_INFO(DispStn).ADF_TANKTYPE > 10 And STN_INFO(DispStn).ADF_TANKTYPE <= 20) Then
                            chkADF_Heater.Enabled = True
                            chkADF_Heater.Visible = True
                            lblADF_HeaterSP.Visible = True
                            If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                            If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                            txtADF_HeaterSP.Enabled = True
                            txtADF_HeaterSP.Visible = True
                            If USINGC Then txtADF_HeaterSP.ToolTipText = "15 to 50 deg C"
                            If USINGF Then txtADF_HeaterSP.ToolTipText = "60 to 120 deg F"
                        Else
                            chkADF_Heater.Visible = False
                            lblADF_HeaterSP.Visible = False
                            txtADF_HeaterSP.Visible = False
                        End If
                        txtButnPercent.Visible = False
                        lblButnPercent.Visible = False
                        lblButnPercentUnits.Visible = False
                        chkLoadRatePID.Visible = True
                    Else
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format(Stn_AIO(DispStn, asNitrogenFlowSP).EuMax, "##0.#") & " slpm"
                        txtButnPercent.Visible = True
                        lblButnPercent.Visible = True
                        lblButnPercentUnits.Visible = True
                        chkLoadRatePID.Visible = False
                    End If
                Case STN_LIVEORVR2_TYPE
                    frmLiveFuel.Visible = True
                    chkOrvrMfc.Visible = True
                    If DspRecipe.LiveFuel Then
                        ' livefuel MFC
                        If DspRecipe.UseHiRangeMFC Then
                            ' hi-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " slpm"
                        Else
                            ' lo-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.#") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " slpm"
                        End If
                        If (STN_INFO(DispStn).ADF_TANKTYPE > 10 And STN_INFO(DispStn).ADF_TANKTYPE <= 20) Then
                            chkADF_Heater.Enabled = True
                            chkADF_Heater.Visible = True
                            lblADF_HeaterSP.Visible = True
                            If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                            If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                            txtADF_HeaterSP.Enabled = True
                            txtADF_HeaterSP.Visible = True
                            If USINGC Then txtADF_HeaterSP.ToolTipText = "15 to 50 deg C"
                            If USINGF Then txtADF_HeaterSP.ToolTipText = "60 to 120 deg F"
                        Else
                            chkADF_Heater.Visible = False
                            lblADF_HeaterSP.Visible = False
                            txtADF_HeaterSP.Visible = False
                        End If
                        txtButnPercent.Visible = False
                        lblButnPercent.Visible = False
                        lblButnPercentUnits.Visible = False
                        chkLoadRatePID.Visible = True
                    Else
                        If DspRecipe.UseHiRangeMFC Then
                            ' hi-range MFC
                            txtButnPercent.Visible = True
                            lblButnPercent.Visible = True
                            lblButnPercentUnits.Visible = True
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " slpm"
                        Else
                            ' lo-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.#") & " slpm"
                            txtButnPercent.Visible = True
                            lblButnPercent.Visible = True
                            lblButnPercentUnits.Visible = True
                            chkLoadRatePID.Visible = False
                        End If
                    End If
                Case STN_COMBO3_TYPE
                    ' future
                Case Else
                    frmLiveFuel.Visible = True
                    chkOrvrMfc.Visible = True
                    txtButnPercent.Visible = True
                    lblButnPercent.Visible = True
                    lblButnPercentUnits.Visible = True
            End Select
            chkLiveFuel.Value = IIf(DspRecipe.LiveFuel, cYES, cNO)
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

Private Sub DspRcpToMemRcp()
    MemRecipe = DspRecipe
End Sub

Private Sub MemRcpToDspRcp()
    DspRecipe = MemRecipe
End Sub

Private Sub ScreenToDspRcp()
' Copies Screen data to DspRecipe
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 5
Dim Idx As Integer

    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    
    DspRecipe.Number = 0
    If IsNumeric(pnlDispRcpNum.Caption) Then DspRecipe.Number = CInt(pnlDispRcpNum.Caption)
    DspRecipe.Name = txtRecipeName.text
    
    DspRecipe.NitrogenFlow = CSng(txtNitrogenFlow.text)
    DspRecipe.NitrogenFlowSave = DspRecipe.NitrogenFlow
    DspRecipe.Load_Rate = CSng(txtLoadRate.text)
    DspRecipe.Load_RateSave = DspRecipe.Load_Rate
    DspRecipe.UseLoadRatePID = IIf((chkLoadRatePID.Value = cYES), True, False)
    DspRecipe.Mix_Percent = CSng(txtButnPercent.text)
    DspRecipe.WC_Mult = CSng(txtWorkCapMult)
    DspRecipe.WC_MultSave = DspRecipe.WC_Mult
    DspRecipe.EPAFill = CSng(txtEPAFill.text)
    DspRecipe.Load_Wt = CSng(txtTargetWt.text)
    DspRecipe.LoadBreakthrough = CSng(txtLoadBreakthrough.text)
'    DspRecipe.FIDmg = CSng(txtFIDmg.text)
    DspRecipe.Load_Time = CSng(txtLoadTime.text)
    
    ' cycle options
    If tabsCycletype.Tabs(CyclePurgeLoad).Selected Then DspRecipe.CycleType = CyclePurgeLoad
    If tabsCycletype.Tabs(CycleLoadPurge).Selected Then DspRecipe.CycleType = CycleLoadPurge
    ' purge method
    If optNoPurge.Value = cYES Then DspRecipe.Purge_Method = NOPURGE
    If optPurgeTime.Value = cYES Then DspRecipe.Purge_Method = PURGEBYTIME
    If optPurgeLiters.Value = cYES Then DspRecipe.Purge_Method = PURGEBYLITERS
    If optPurgeVolume.Value = cYES Then DspRecipe.Purge_Method = PURGEBYVOLUME
    If optPurgeAuxOnly.Value = cYES Then DspRecipe.Purge_Method = PURGEAUXONLY
    If optPurgeProfile.Value = cYES Then DspRecipe.Purge_Method = PURGEBYPROFILE
    If optPurgeWC.Value = cYES Then DspRecipe.Purge_Method = PURGEBYWC
    If optPurgeTarget.Value = cYES Then DspRecipe.Purge_Method = PURGETOTARGET
    If optPurgeUndo.Value = cYES Then DspRecipe.Purge_Method = PURGETOUNDOLOAD
    ' purge-to-target mode
    If optTargetContinuous.Value Then DspRecipe.Purge_TargetMode = TARGETCONTINUOUS
    If optTargetPurgePauseRepeat.Value Then DspRecipe.Purge_TargetMode = TARGETPURGEPAUSE
    ' purge options
    DspRecipe.Purge_AuxTime = CSng(txtPurgeAuxOnly.text)
    DspRecipe.Purge_Time = CSng(txtPurgeTime.text)
    DspRecipe.Purge_Flow = CSng(txtPurgeFlow.text)
    DspRecipe.Purge_Liters = CSng(txtPurgeLiters.text)
    DspRecipe.Purge_Can_Vol = CSng(txtPurgeVolume.text)
    DspRecipe.Purge_ProfileNumber = CInt(ValueFromText(Mid(txtPurgeProfile.text, 2, (Len(txtPurgeProfile.text) - 1))))
    DspRecipe.Purge_TargetWC = CSng(txtPurgeWC.text)
    DspRecipe.Purge_TargetWeight = CSng(txtPurgeTarget.text)
    DspRecipe.Purge_MaxVolumes = CInt(txtTargetTimeout.text)
    DspRecipe.Purge_TargetPurge = CSng(txtTargetPurge.text)
    DspRecipe.Purge_TargetPause = CSng(txtTargetPause.text)
    
    DspRecipe.PurgeAuxCan = IIf((chkPurgeAuxCan.Value = cYES), True, False)
    DspRecipe.PurgeCansInSeries = IIf((chkPurgeCansInSeries.Value = cYES), True, False)
    DspRecipe.PurgeOven = IIf((chkUsePurgeOven.Value = cYES), True, False)
    DspRecipe.PurgeOvenSP = CSng(txtPurgeOvenSP.text)
    DspRecipe.UseAuxScale = IIf((chkUseAuxScale.Value = cYES), True, False)
    DspRecipe.AuxScaleNo = CInt(txtAuxScaleNo.text)
    DspRecipe.PauseLeakTime = CSng(txtPauseLeakTime.text)
    DspRecipe.PauseLoadTime = CSng(txtPauseLoadTime.text)
    DspRecipe.PausePurgeTime = CSng(txtPausePurgeTime.text)
    DspRecipe.UsePriScale = IIf((chkPrimaryScale.Value = cYES), True, False)
    DspRecipe.PriScaleNo = CInt(txtPrimaryScaleNo.text)
    DspRecipe.PauseAfterLeak = IIf((chkPauseAfterLeak.Value = cYES), True, False)
    DspRecipe.PauseAfterLoad = IIf(optPauseAfterLoad.Value, True, False)
    DspRecipe.PauseAfterLoadForOper = IIf(optPauseAfterLoadForOper.Value, True, False)
    DspRecipe.PauseAfterPurge = IIf(optPauseAfterPurge.Value, True, False)
    DspRecipe.PauseAfterPurgeForOper = IIf(optPauseAfterPurgeForOper.Value, True, False)
'    DspRecipe.TargetConcentration = frmAnalyzer.txtTargetConcentration
'    DspRecipe.DwellTime = CSng(frmAnalyzer.txtDwellTime.text)
    DspRecipe.LeakCheck = IIf((chkLeakCheck.Value = cYES), True, False)
    DspRecipe.LeakPrimary = IIf((chkLeakPrimary.Value = cYES), True, False)
    DspRecipe.LeakAux = IIf((USINGAUXLEAKCHECK And (chkLeakAux.Value = cYES)), True, False)
'    DspRecipe.UseAnalyzer = IIf((chkUseAnalyzer.Value = cYES), True, False)
    DspRecipe.MaxLoadTime = txtMaxLoadTime
    DspRecipe.UseHiRangeMFC = IIf((chkOrvrMfc.Value = cYES), True, False)
    
    DspRecipe.IDLoad = CSng(txtIDLoad.text)
    DspRecipe.LoadL = CSng(txtLoadL.text)
    DspRecipe.LoadV = CSng(txtLoadV.text)
    DspRecipe.IDPurge = CSng(txtIDPurge.text)
    DspRecipe.PurgeL = CSng(txtPurgeL.text)
    DspRecipe.PurgeV = CSng(txtPurgeV.text)
    DspRecipe.IDVent = CSng(txtIDVent.text)
    DspRecipe.VentL = CSng(txtVentL.text)
    DspRecipe.VentV = CSng(txtVentV.text)
    
    DspRecipe.LiveFuel = IIf((chkLiveFuel.Value = cYES), True, False)
    DspRecipe.LiveFuelChgAuto = IIf((chkLiveFuelChgAuto.Value = cYES), True, False)
    DspRecipe.LiveFuelChgFreq = CInt(txtLiveFuelChgFreq.text)
    DspRecipe.ADF_Heater = IIf((chkADF_Heater.Value = cYES), True, False)
    DspRecipe.ADF_HeaterSP = CSng(txtADF_HeaterSP.text)
    
    ' start method
    If optStartNow.Value = cYES Then
        DspRecipe.StartMethod = STARTNOW
        DspRecipe.StartDelay = 0
        DspRecipe.StartDate = Now()
    ElseIf optStartAfter.Value = cYES Then
        DspRecipe.StartMethod = STARTDELAYED
        DspRecipe.StartDelay = CDbl(txtStartAfterMin.text)
        DspRecipe.StartDate = Now()
    ElseIf optStartAt.Value = cYES Then
        DspRecipe.StartMethod = STARTATDATE
        DspRecipe.StartDelay = 0
        DspRecipe.StartDate = CDate(txtStartAtDate.text)
    Else
        DspRecipe.StartMethod = STARTNOW
        DspRecipe.StartDelay = 0
        DspRecipe.StartDate = Now()
    End If
  
    ' end method
    If optEndCycles.Value = cYES Then
        DspRecipe.EndMethod = ENDCYCLES
    ElseIf optEndWeightChange.Value = cYES Then
        DspRecipe.EndMethod = ENDWEIGHTCHG
    Else
        DspRecipe.EndMethod = ENDCYCLES
    End If
    DspRecipe.UpdateCanWc = IIf((optUpdateCanWc.Value = cYES), True, False)
    DspRecipe.Cycles = CInt(txtPFCycle)
    DspRecipe.CyclesSave = DspRecipe.Cycles
    DspRecipe.EndWeightTolerance = CSng(txtWeightChangeTol.text)
    DspRecipe.EndConsecutiveCycles = CInt(txtConsecutiveCycles.text)
    DspRecipe.EndMaximumCycles = CInt(txtMaximumCycles.text)
    DspRecipe.EndMinimumCycles = CInt(txtMinimumCycles.text)
    
    ' load method
    If optNoLoad.Value = cYES Then DspRecipe.Load_Method = NOLOAD
    If optLoadTime.Value = cYES Then DspRecipe.Load_Method = LOADBYTIME
    If optWcm.Value = cYES Then DspRecipe.Load_Method = LOADBYWC
    If optLoadweight.Value = cYES Then DspRecipe.Load_Method = LOADBYWEIGHT
    If optLoadBreakthrough.Value = cYES Then DspRecipe.Load_Method = LOADBYBREAKTHRU
    If optFIDBreakthrough.Value = cYES Then DspRecipe.Load_Method = LOADBYFID
    DspRecipe.Load_MethodSave = DspRecipe.Load_Method

    ' aux outputs
    DspRecipe.AuxOutputs = IIf((chkAuxOutputs.Value = cON), True, False)
    For Idx = 1 To 4
        DspRecipe.AuxOutputs_Load(Idx) = IIf((chkAuxLoad(Idx).Value = cON), True, False)
        DspRecipe.AuxOutputs_Purge(Idx) = IIf((chkAuxPurge(Idx).Value = cON), True, False)
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

Public Sub ExitScreen()
    ' close canister / recipe database
    dbDbase.Close
    ' unload form
    frmRecipe.Visible = False
    Unload Me
End Sub

Public Sub RecipeDisplay_ByNum()
    GetRecipe MASTERMODE, DispRcp, 0
    DspRcpToScreen
    Chgs = False
End Sub

Public Sub RecipeDisplay_ByStnShift()
    GetRecipe STATIONMODE, DispStn, DispShift
    DspRcpToScreen
    Chgs = False
End Sub

Private Sub GetRecipe(ByVal MstStnMode As Integer, ByVal index1 As Integer, ByVal index2 As Integer)
    Select Case MstStnMode
        Case MASTERMODE
            ' Read Master Recipe Record
            Criteria = "SELECT * FROM [MasterRecipe] WHERE [Number] = " & index1 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
        Case STATIONMODE
            ' Read Station Recipe Record
            Criteria = "SELECT * FROM [StationRecipe] WHERE [Station] = " & index1 & "  and [Shift] = " & index2 & " "
            Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
            
    End Select
    ' was a record found ??
    If rsRecord.BOF Then
        ' no record; display blank recipe
        InitDspRcp MstStnMode, index1, index2
    Else
        ' record found; display it
        DbToDspRcp
    End If
    ' done
    rsRecord.Close
End Sub

Private Sub DbToDspRcp()
        
    ' Load Recipe Record to DspRecipe
    DispRcp = rsRecord("Number")
    DspRecipe.Number = rsRecord("Number")
    DspRecipe.Name = rsRecord("Name")
   
    DspRecipe.CycleType = rsRecord("CycleType")
    If (DspRecipe.CycleType = CycleUndefined) Then DspRecipe.CycleType = CyclePurgeLoad
       
    DspRecipe.Load_Method = rsRecord("Load_Method")
    DspRecipe.Load_MethodSave = DspRecipe.Load_Method
    DspRecipe.NitrogenFlow = rsRecord("NitrogenFlow")
    DspRecipe.NitrogenFlowSave = DspRecipe.NitrogenFlow
    DspRecipe.Load_Rate = rsRecord("Load_Rate")
    DspRecipe.Load_RateSave = DspRecipe.Load_Rate
    DspRecipe.Mix_Percent = rsRecord("Mix_Percent")
    DspRecipe.WC_Mult = rsRecord("WC_Mult")
    DspRecipe.WC_MultSave = DspRecipe.WC_Mult
    DspRecipe.EPAFill = rsRecord("EPAFill")
    DspRecipe.Load_Wt = rsRecord("Load_Wt")
    DspRecipe.LoadBreakthrough = rsRecord("LoadBreakthrough")
'    DspRecipe.FIDmg = rsRecord("FIDmg")
    DspRecipe.Load_Time = rsRecord("Load_Time")
    If IsNumeric(rsRecord("Purge_Method")) Then
        DspRecipe.Purge_Method = rsRecord("Purge_Method")
    Else
        DspRecipe.Purge_Method = NOPURGE
    End If
    If IsNumeric(rsRecord("Purge_TargetMode")) Then
        DspRecipe.Purge_TargetMode = rsRecord("Purge_TargetMode")
    Else
        DspRecipe.Purge_TargetMode = NOTARGET
    End If
    If IsNumeric(rsRecord("Purge_AuxTime")) Then
        DspRecipe.Purge_AuxTime = rsRecord("Purge_AuxTime")
    Else
        DspRecipe.Purge_AuxTime = 0
    End If
    DspRecipe.Purge_Time = rsRecord("Purge_Time")
    DspRecipe.Purge_Flow = rsRecord("Purge_Flow")
    DspRecipe.Purge_Liters = rsRecord("Purge_Liters")
    DspRecipe.Purge_Can_Vol = rsRecord("Purge_Can_Vol")
    DspRecipe.Purge_ProfileNumber = rsRecord("Purge_ProfileNumber")
    DspRecipe.Purge_TargetWC = rsRecord("Purge_TargetWC")
    DspRecipe.Purge_TargetWeight = rsRecord("Purge_TargetWeight")
    DspRecipe.Purge_MaxVolumes = rsRecord("Purge_MaxVolumes")
    DspRecipe.Purge_TargetPurge = rsRecord("Purge_TargetPurge")
    DspRecipe.Purge_TargetPause = rsRecord("Purge_TargetPause")
    
    DspRecipe.PurgeAuxCan = rsRecord("PurgeAuxCan")
    DspRecipe.PurgeCansInSeries = rsRecord("PurgeCansInSeries")
    DspRecipe.PurgeOven = rsRecord("PurgeInOven")
    DspRecipe.PurgeOvenSP = rsRecord("PurgeOvenSP")
    DspRecipe.UseAuxScale = rsRecord("UseAuxScale")
    DspRecipe.AuxScaleNo = rsRecord("AuxScaleNo")
    DspRecipe.PauseLeakTime = rsRecord("PauseLeakTime")
    DspRecipe.PauseLoadTime = rsRecord("PauseLoadTime")
    DspRecipe.PausePurgeTime = rsRecord("PausePurgeTime")
    DspRecipe.UsePriScale = rsRecord("UsePriScale")
    DspRecipe.PriScaleNo = rsRecord("PriScaleNo")
    DspRecipe.PauseAfterLeak = rsRecord("PauseAfterLeak")
    DspRecipe.PauseAfterLoad = rsRecord("PauseAfterLoad")
    DspRecipe.PauseAfterLoadForOper = rsRecord("PauseAfterLoadForOper")
    DspRecipe.PauseAfterPurge = rsRecord("PauseAfterPurge")
    DspRecipe.PauseAfterPurgeForOper = rsRecord("PauseAfterPurgeForOper")
'    DspRecipe.TargetConcentration = rsRecord("TargetConcentration")
'    DspRecipe.DwellTime = rsRecord("DwellTime")
    DspRecipe.LeakCheck = rsRecord("LeakCheck")
    DspRecipe.LeakPrimary = rsRecord("LeakPrimary")
    DspRecipe.LeakAux = rsRecord("LeakAux")
'    DspRecipe.UseAnalyzer = rsRecord("UseAnalyzer")
    DspRecipe.MaxLoadTime = rsRecord("MaxLoadTime")
    DspRecipe.UseHiRangeMFC = rsRecord("UseHiRangeMFC")
    DspRecipe.UseLoadRatePID = rsRecord("UseLoadRatePID")
    
    DspRecipe.IDLoad = rsRecord("IDLoad")
    DspRecipe.LoadL = rsRecord("LoadL")
    DspRecipe.LoadV = rsRecord("LoadV")
    DspRecipe.IDPurge = rsRecord("IDPurge")
    DspRecipe.PurgeL = rsRecord("PurgeL")
    DspRecipe.PurgeV = rsRecord("PurgeV")
    DspRecipe.IDVent = rsRecord("IDVent")
    DspRecipe.VentL = rsRecord("VentL")
    DspRecipe.VentV = rsRecord("VentV")
    
    DspRecipe.LiveFuel = rsRecord("LiveFuel")
    DspRecipe.LiveFuelChgAuto = rsRecord("LiveFuelChgAuto")
    DspRecipe.LiveFuelChgFreq = rsRecord("LiveFuelChgFreq")
    DspRecipe.ADF_Heater = rsRecord("ADF_Heater")
    DspRecipe.ADF_HeaterSP = rsRecord("ADF_HeaterSP")
    
    ' start method
    DspRecipe.StartMethod = rsRecord("StartMethod")
    DspRecipe.StartDelay = rsRecord("StartDelay")
    DspRecipe.StartDate = rsRecord("StartDate")
    
    ' end method
    DspRecipe.EndMethod = rsRecord("EndMethod")
    DspRecipe.EndWeightTolerance = rsRecord("EndWeightTolerance")
    DspRecipe.EndMaximumCycles = rsRecord("EndMaximumCycles")
    DspRecipe.EndMinimumCycles = rsRecord("EndMinimumCycles")
    DspRecipe.EndConsecutiveCycles = rsRecord("EndConsecutiveCycles")
    DspRecipe.UpdateCanWc = rsRecord("UpdateCanWc")
    DspRecipe.Cycles = rsRecord("Cycles")
    DspRecipe.CyclesSave = DspRecipe.Cycles
    
    ' aux outputs
    DspRecipe.AuxOutputs = rsRecord("AuxOutputs")
    DspRecipe.AuxOutputs_Load(1) = rsRecord("AuxOutput1_Load")
    DspRecipe.AuxOutputs_Purge(1) = rsRecord("AuxOutput1_Purge")
    DspRecipe.AuxOutputs_Load(2) = rsRecord("AuxOutput2_Load")
    DspRecipe.AuxOutputs_Purge(2) = rsRecord("AuxOutput2_Purge")
    DspRecipe.AuxOutputs_Load(3) = rsRecord("AuxOutput3_Load")
    DspRecipe.AuxOutputs_Purge(3) = rsRecord("AuxOutput3_Purge")
    DspRecipe.AuxOutputs_Load(4) = rsRecord("AuxOutput4_Load")
    DspRecipe.AuxOutputs_Purge(4) = rsRecord("AuxOutput4_Purge")
    
    If RecipeMode = STATIONMODE Then
        If USINGHARDPIPEDSCALES Then
            DspRecipe.AuxScaleNo = STN_INFO(DispStn).DefAuxScale
            DspRecipe.PriScaleNo = STN_INFO(DispStn).DefPriScale
        Else
            If DspRecipe.UseAuxScale And DspRecipe.AuxScaleNo = 0 Then DspRecipe.AuxScaleNo = DispStn
            If DspRecipe.UsePriScale And DspRecipe.PriScaleNo = 0 Then DspRecipe.PriScaleNo = DispStn
        End If
    End If

        
End Sub

Private Sub SaveMasterRcp(ByVal index1 As Integer)
        
    ' Read Master Recipe Record
    Criteria = "SELECT * FROM [MasterRecipe] WHERE [Number] = " & index1 & " "
    Set rsRecord = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
    
    If rsRecord.BOF Then
        rsRecord.AddNew
        rsRecord("Number") = index1
    Else
      rsRecord.MoveFirst
      rsRecord.Edit
    End If
       
    ' Update Master Recipe Record
    rsRecord("Name") = DspRecipe.Name
   
    rsRecord("CycleType") = DspRecipe.CycleType
    rsRecord("CycleTypeDesc") = CycleTypeDesc(DspRecipe.CycleType)
            
    rsRecord("Load_Method") = DspRecipe.Load_Method
    rsRecord("Load_MethodDesc") = LoadMethodDesc(DspRecipe.Load_Method)
    rsRecord("NitrogenFlow") = DspRecipe.NitrogenFlow
    rsRecord("Load_Rate") = DspRecipe.Load_Rate
    rsRecord("Mix_Percent") = DspRecipe.Mix_Percent
    rsRecord("WC_Mult") = DspRecipe.WC_Mult
    rsRecord("EPAFill") = DspRecipe.EPAFill
    rsRecord("Load_Wt") = DspRecipe.Load_Wt
    rsRecord("LoadBreakthrough") = DspRecipe.LoadBreakthrough
    rsRecord("Load_Time") = DspRecipe.Load_Time
    rsRecord("Purge_Method") = DspRecipe.Purge_Method
    rsRecord("Purge_MethodDesc") = PurgeMethodDesc(DspRecipe.Purge_Method)
    rsRecord("Purge_AuxTime") = DspRecipe.Purge_AuxTime
    rsRecord("Purge_Time") = DspRecipe.Purge_Time
    rsRecord("Purge_Flow") = DspRecipe.Purge_Flow
    rsRecord("Purge_Liters") = DspRecipe.Purge_Liters
    rsRecord("Purge_Can_Vol") = DspRecipe.Purge_Can_Vol
    rsRecord("Purge_ProfileNumber") = DspRecipe.Purge_ProfileNumber
    rsRecord("Purge_TargetMode") = DspRecipe.Purge_TargetMode
    rsRecord("Purge_TargetModeDesc") = PurgeTargetDesc(DspRecipe.Purge_TargetMode)
    rsRecord("Purge_TargetWC") = DspRecipe.Purge_TargetWC
    rsRecord("Purge_TargetWeight") = DspRecipe.Purge_TargetWeight
    rsRecord("Purge_MaxVolumes") = DspRecipe.Purge_MaxVolumes
    rsRecord("Purge_TargetPurge") = DspRecipe.Purge_TargetPurge
    rsRecord("Purge_TargetPause") = DspRecipe.Purge_TargetPause
    
    rsRecord("PurgeAuxCan") = DspRecipe.PurgeAuxCan
    rsRecord("PurgeCansInSeries") = DspRecipe.PurgeCansInSeries
    rsRecord("PurgeInOven") = DspRecipe.PurgeOven
    rsRecord("PurgeOvenSP") = DspRecipe.PurgeOvenSP
    rsRecord("UseAuxScale") = DspRecipe.UseAuxScale
    rsRecord("AuxScaleNo") = DspRecipe.AuxScaleNo
    rsRecord("PauseLeakTime") = DspRecipe.PauseLeakTime
    rsRecord("PauseLoadTime") = DspRecipe.PauseLoadTime
    rsRecord("PausePurgeTime") = DspRecipe.PausePurgeTime
    rsRecord("UsePriScale") = DspRecipe.UsePriScale
    rsRecord("PriScaleNo") = DspRecipe.PriScaleNo
    rsRecord("PauseAfterLeak") = DspRecipe.PauseAfterLeak
    rsRecord("PauseAfterLoad") = DspRecipe.PauseAfterLoad
    rsRecord("PauseAfterLoadForOper") = DspRecipe.PauseAfterLoadForOper
    rsRecord("PauseAfterPurge") = DspRecipe.PauseAfterPurge
    rsRecord("PauseAfterPurgeForOper") = DspRecipe.PauseAfterPurgeForOper
'    rsRecord("TargetConcentration") = DspRecipe.TargetConcentration
'    rsRecord("DwellTime") = DspRecipe.DwellTime
    rsRecord("LeakCheck") = DspRecipe.LeakCheck
    rsRecord("LeakPrimary") = DspRecipe.LeakPrimary
    rsRecord("LeakAux") = DspRecipe.LeakAux
'    rsRecord("UseAnalyzer") = DspRecipe.UseAnalyzer
    rsRecord("MaxLoadTime") = DspRecipe.MaxLoadTime
    rsRecord("UseHiRangeMFC") = DspRecipe.UseHiRangeMFC
    rsRecord("UseLoadRatePID") = DspRecipe.UseLoadRatePID
    
    rsRecord("IDLoad") = DspRecipe.IDLoad
    rsRecord("LoadL") = DspRecipe.LoadL
    rsRecord("LoadV") = DspRecipe.LoadV
    rsRecord("IDPurge") = DspRecipe.IDPurge
    rsRecord("PurgeL") = DspRecipe.PurgeL
    rsRecord("PurgeV") = DspRecipe.PurgeV
    rsRecord("IDVent") = DspRecipe.IDVent
    rsRecord("VentL") = DspRecipe.VentL
    rsRecord("VentV") = DspRecipe.VentV
    
    rsRecord("LiveFuel") = DspRecipe.LiveFuel
    rsRecord("LiveFuelChgAuto") = DspRecipe.LiveFuelChgAuto
    rsRecord("LiveFuelChgFreq") = DspRecipe.LiveFuelChgFreq
    rsRecord("ADF_Heater") = DspRecipe.ADF_Heater
    rsRecord("ADF_HeaterSP") = DspRecipe.ADF_HeaterSP
    
    ' start method
    rsRecord("StartMethod") = DspRecipe.StartMethod
    rsRecord("StartMethodDesc") = StartMethodDesc(DspRecipe.StartMethod)
    rsRecord("StartDelay") = DspRecipe.StartDelay
    rsRecord("StartDate") = DspRecipe.StartDate
        
    ' end method
    rsRecord("EndMethod") = DspRecipe.EndMethod
    rsRecord("EndMethodDesc") = EndMethodDesc(DspRecipe.EndMethod)
    rsRecord("EndMaximumCycles") = DspRecipe.EndMaximumCycles
    rsRecord("EndMinimumCycles") = DspRecipe.EndMinimumCycles
    rsRecord("EndConsecutiveCycles") = DspRecipe.EndConsecutiveCycles
    rsRecord("EndWeightTolerance") = DspRecipe.EndWeightTolerance
    rsRecord("UpdateCanWc") = DspRecipe.UpdateCanWc
    rsRecord("Cycles") = DspRecipe.Cycles
        
    ' aux outputs
    rsRecord("AuxOutputs") = DspRecipe.AuxOutputs
    rsRecord("AuxOutput1_Load") = DspRecipe.AuxOutputs_Load(1)
    rsRecord("AuxOutput1_Purge") = DspRecipe.AuxOutputs_Purge(1)
    rsRecord("AuxOutput2_Load") = DspRecipe.AuxOutputs_Load(2)
    rsRecord("AuxOutput2_Purge") = DspRecipe.AuxOutputs_Purge(2)
    rsRecord("AuxOutput3_Load") = DspRecipe.AuxOutputs_Load(3)
    rsRecord("AuxOutput3_Purge") = DspRecipe.AuxOutputs_Purge(3)
    rsRecord("AuxOutput4_Load") = DspRecipe.AuxOutputs_Load(4)
    rsRecord("AuxOutput4_Purge") = DspRecipe.AuxOutputs_Purge(4)
    
    rsRecord.Update
    rsRecord.Close

End Sub

Public Sub InitDspRcp(ByVal MstStnMode As Integer, ByVal index1 As Integer, ByVal index2 As Integer)
' Initializes DspRecipe
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 1

    Select Case MstStnMode
        Case MASTERMODE
            ' master
            DspRecipe = EmptyRecipe
            DspRecipe.Number = CInt(index1)
            
        Case STATIONMODE
            ' station
            DspRecipe = EmptyRecipe
            DspRecipe.Number = CInt(0)
    
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

Public Sub InitRecipe()
    Select Case RecipeMode
        Case MASTERMODE
            ' master
            If (DispRcp < 1 Or DispRcp > NR_RCP) Then DispRcp = 1
            GetRecipe MASTERMODE, DispRcp, 0
            txtPurgeProfile.Visible = True
            cmdPurgeProfile.ToolTipText = "Select a Purge Profile"
            cmdRestore.Visible = False
            cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
            cmdPrint.ToolTipText = "Print Recipe"
            cmdPrintAll.Visible = IIf(PRINTERAVAILABLE, True, False)
            cmdPrintAll.ToolTipText = "Print All Master Recipes"
        Case STATIONMODE
            ' station
            If StationRecipe(DispStn, DispShift).Number < 0 _
             Or StationRecipe(DispStn, DispShift).Number > NR_RCP Then
               DispRcp = 0
            Else
               DispRcp = StationRecipe(DispStn, DispShift).Number
            End If
            GetRecipe STATIONMODE, DispStn, DispShift
            If StationControl(DispStn, DispShift).Mode <> VBIDLE Then
               cmdRestore.Visible = False
               cmdSave.Visible = False
            Else
               cmdRestore.Visible = True
               cmdSave.Visible = True
            End If
            txtPurgeProfile.Visible = False
            cmdPurgeProfile.ToolTipText = "Open the Purge Profile Configuration screen"
            cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
            cmdPrint.ToolTipText = "Print Recipe"
            cmdPrintAll.Visible = IIf(PRINTERAVAILABLE, True, False)
            cmdPrintAll.ToolTipText = "Print All Station Recipes"
    End Select
    DspRcpToScreen
    Chgs = False
    tmrUpdate.Enabled = True
End Sub

Private Sub UpdateRecipe()
    If (DispRcp < 1 Or DispRcp > NR_RCP) Then DispRcp = 1
    GetRecipe MASTERMODE, DispRcp, 0
    DspRcpToScreen
    Chgs = False
End Sub

Private Sub chkAuxOutputs_Click()
    ' aux outputs selection button
    cmdCfgAuxOutputs.Visible = IIf(((USING_AUX_OUTPUTS) And (chkAuxOutputs.Value = cON)), True, False)
End Sub

Private Sub chkLoadRatePID_Click()
    chkLoadRatePID.BackColor = frmNotHighlight.BackColor
    Select Case chkLoadRatePID.Value
        Case cYES
            txtNitrogenFlow.text = "0.0"
            txtNitrogenFlow.Enabled = False
        Case cNO
            txtNitrogenFlow.Enabled = True
    End Select
End Sub

Private Sub chkOrvrMfc_Click()
    chkOrvrMfc.BackColor = frmNotHighlight.BackColor
    Select Case RecipeMode
        Case MASTERMODE
            txtNitrogenFlow.ToolTipText = "0.1 to 50 slpm"
        Case STATIONMODE
            Select Case STN_INFO(DispStn).Type
                Case STN_REGULAR_TYPE
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " slpm"
                Case STN_ORVR_TYPE
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " slpm"
                Case STN_ORVR2_TYPE
                    If DspRecipe.UseHiRangeMFC Then
                        ' hi-range MFC
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " slpm"
                    Else
                        ' lo-range MFC
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " slpm"
                    End If
                Case STN_LIVEFUEL_TYPE
                    txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.#") & " slpm"
                Case STN_LIVEREG_TYPE
                    If DspRecipe.LiveFuel Then
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.#") & " slpm"
                    Else
                        txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.#") & " slpm"
                    End If
                Case STN_LIVEORVR2_TYPE
                    If DspRecipe.LiveFuel Then
                        ' livefuel MFC
                        If DspRecipe.UseHiRangeMFC Then
                            ' hi-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " slpm"
                        Else
                            ' lo-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "##0.#") & " to " & Format((0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "##0.#") & " slpm"
                        End If
                    Else
                        If DspRecipe.UseHiRangeMFC Then
                            ' hi-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "##0.0") & " slpm"
                        Else
                            ' lo-range MFC
                            txtNitrogenFlow.ToolTipText = Format((0.05 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.0") & " to " & Format((0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "##0.#") & " slpm"
                        End If
                    End If
                Case STN_COMBO3_TYPE
                    ' future
                Case Else
                    txtNitrogenFlow.ToolTipText = "0.1 to 50 slpm"
            End Select
    End Select
End Sub

Private Sub chkPurgeCansInSeries_Click()
    If (chkPurgeCansInSeries.Value = cYES) Then
        chkPurgeAuxCan.Value = cYES
    End If
End Sub

Private Sub chkUsePurgeOven_Click()
    If (chkUsePurgeOven.Value = cYES) Then
        optPurgeWC.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtPurgeWC.Enabled = False
        txtPurgeTarget.Enabled = False
        optPurgeWC.Enabled = False
        optPurgeTarget.Enabled = False
        optPurgeUndo.Enabled = False
    Else
        txtPurgeWC.Enabled = True
        txtPurgeTarget.Enabled = True
        optPurgeWC.Enabled = True
        optPurgeTarget.Enabled = True
        optPurgeUndo.Enabled = True
        txtPurgeOvenSP.BackColor = txtNotHighlight.BackColor
    End If
End Sub

Private Sub cmdCfgAuxOutputs_Click()
    ' aux outputs
    frmAuxOutputs.Top = IIf((frmAuxOutputs.Top = frmStatus.Top), OutOfSight, frmStatus.Top)
End Sub

Private Sub cmdClose_Click()
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
End Sub

Private Sub cmdCloseInfo_Click()
    frmRcpInfo.Top = OutOfSight
End Sub

Private Sub cmdCopy_Click()
    ScreenToDspRcp
    DspRcpToMemRcp
End Sub

Private Sub cmdEndMethodInfo_Click()
    frmRcpInfo.Top = frmEnd.Top
End Sub

Private Sub cmdOpen_Click()
    ' close aux outputs
    frmAuxOutputs.Top = OutOfSight
    ' open Master Recipes selection screen
    frmSearchRcp.Show
    frmSearchRcp.ChgRecipeMode RecipeMode
    frmSearchRcp.ChgSelectionDestination rcpdestRecipe
End Sub

Private Sub cmdPaste_Click()
    MemRcpToDspRcp
    DspRecipe.Number = DispRcp
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        DspRecipe.AuxScaleNo = STN_INFO(DispStn).DefAuxScale
        DspRecipe.PriScaleNo = STN_INFO(DispStn).DefPriScale
    End If
    DspRcpToScreen
    Chgs = True
End Sub

Private Sub Reset_BackColors()
'
' resets the background colors
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 6

    txtRecipeName.BackColor = txtNotHighlight.BackColor
    
    optPauseAfterPurge.BackColor = frmNotHighlight.BackColor
    optPauseAfterPurgeForOper.BackColor = frmNotHighlight.BackColor
    chkPrimaryScale.BackColor = frmNotHighlight.BackColor
    chkUseAuxScale.BackColor = frmNotHighlight.BackColor
    optPauseAfterLoad.BackColor = frmNotHighlight.BackColor
    optPauseAfterLoadForOper.BackColor = frmNotHighlight.BackColor
    chkPurgeAuxCan.BackColor = frmNotHighlight.BackColor
    chkLeakCheck.BackColor = frmNotHighlight.BackColor
    chkPauseAfterLeak.BackColor = frmNotHighlight.BackColor
    chkLeakPrimary.BackColor = frmNotHighlight.BackColor
    chkLeakAux.BackColor = frmNotHighlight.BackColor
    chkLiveFuel.BackColor = frmNotHighlight.BackColor
    chkLiveFuelChgAuto.BackColor = frmNotHighlight.BackColor
    
    ' START METHOD
    If optStartNow.Value = cYES Then
        optStartNow.BackColor = frmHighlight.BackColor
        optStartNow.ForeColor = frmHighlight.ForeColor
    Else
        optStartNow.BackColor = frmNotHighlight.BackColor
        optStartNow.ForeColor = frmNotHighlight.ForeColor
    End If
    If optStartAfter.Value = cYES Then
        optStartAfter.BackColor = frmHighlight.BackColor
        optStartAfter.ForeColor = frmHighlight.ForeColor
        txtStartAfterMin.BackColor = frmHighlight.ForeColor
        txtStartAfterMin.ForeColor = frmHighlight.BackColor
    Else
        optStartAfter.BackColor = frmNotHighlight.BackColor
        optStartAfter.ForeColor = frmNotHighlight.ForeColor
        txtStartAfterMin.BackColor = txtNotHighlight.BackColor
        txtStartAfterMin.ForeColor = txtNotHighlight.ForeColor
    End If
    If optStartAt.Value = cYES Then
        optStartAt.BackColor = frmHighlight.BackColor
        optStartAt.ForeColor = frmHighlight.ForeColor
        txtStartAtDate.BackColor = frmHighlight.ForeColor
        txtStartAtDate.ForeColor = frmHighlight.BackColor
    Else
        optStartAt.BackColor = frmNotHighlight.BackColor
        optStartAt.ForeColor = frmNotHighlight.ForeColor
        txtStartAtDate.BackColor = txtNotHighlight.BackColor
        txtStartAtDate.ForeColor = txtNotHighlight.ForeColor
    End If
    
    ' PURGE METHOD
    If optNoPurge.Value = cYES Then
        optNoPurge.BackColor = frmHighlight.BackColor
        optNoPurge.ForeColor = frmHighlight.ForeColor
    Else
        optNoPurge.BackColor = frmNotHighlight.BackColor
        optNoPurge.ForeColor = frmNotHighlight.ForeColor
    End If
    If optPurgeTime.Value = cYES Then
        optPurgeTime.BackColor = frmHighlight.BackColor
        optPurgeTime.ForeColor = frmHighlight.ForeColor
        txtPurgeTime.BackColor = frmHighlight.ForeColor
        txtPurgeTime.ForeColor = frmHighlight.BackColor
    Else
        optPurgeTime.BackColor = frmNotHighlight.BackColor
        optPurgeTime.ForeColor = frmNotHighlight.ForeColor
        txtPurgeTime.BackColor = txtNotHighlight.BackColor
        txtPurgeTime.ForeColor = txtNotHighlight.ForeColor
    End If
    If optPurgeVolume.Value = cYES Then
        optPurgeVolume.BackColor = frmHighlight.BackColor
        optPurgeVolume.ForeColor = frmHighlight.ForeColor
        txtPurgeVolume.BackColor = frmHighlight.ForeColor
        txtPurgeVolume.ForeColor = frmHighlight.BackColor
    Else
        optPurgeVolume.BackColor = frmNotHighlight.BackColor
        optPurgeVolume.ForeColor = frmNotHighlight.ForeColor
        txtPurgeVolume.BackColor = txtNotHighlight.BackColor
        txtPurgeVolume.ForeColor = txtNotHighlight.ForeColor
    End If
    If optPurgeAuxOnly.Value = cYES Then
        optPurgeAuxOnly.BackColor = frmHighlight.BackColor
        optPurgeAuxOnly.ForeColor = frmHighlight.ForeColor
        txtPurgeAuxOnly.BackColor = frmHighlight.ForeColor
        txtPurgeAuxOnly.ForeColor = frmHighlight.BackColor
    Else
        optPurgeAuxOnly.BackColor = frmNotHighlight.BackColor
        optPurgeAuxOnly.ForeColor = frmNotHighlight.ForeColor
        txtPurgeAuxOnly.BackColor = txtNotHighlight.BackColor
        txtPurgeAuxOnly.ForeColor = txtNotHighlight.ForeColor
    End If
    If optPurgeProfile.Value = cYES Then
        optPurgeProfile.BackColor = frmHighlight.BackColor
        optPurgeProfile.ForeColor = frmHighlight.ForeColor
        txtPurgeProfile.BackColor = frmHighlight.ForeColor
        txtPurgeProfile.ForeColor = frmHighlight.BackColor
    Else
        optPurgeProfile.BackColor = frmNotHighlight.BackColor
        optPurgeProfile.ForeColor = frmNotHighlight.ForeColor
        txtPurgeProfile.BackColor = frmPurge.BackColor
        txtPurgeProfile.ForeColor = frmPurge.BackColor
    End If
    If optPurgeWC.Value = cYES Then
        optPurgeWC.BackColor = frmHighlight.BackColor
        optPurgeWC.ForeColor = frmHighlight.ForeColor
        txtPurgeWC.BackColor = frmHighlight.ForeColor
        txtPurgeWC.ForeColor = frmHighlight.BackColor
        txtTargetTimeout.BackColor = frmHighlight.ForeColor
        txtTargetTimeout.ForeColor = frmHighlight.BackColor
    Else
        optPurgeWC.BackColor = frmNotHighlight.BackColor
        optPurgeWC.ForeColor = frmNotHighlight.ForeColor
        txtPurgeWC.BackColor = txtNotHighlight.BackColor
        txtPurgeWC.ForeColor = txtNotHighlight.ForeColor
        txtTargetTimeout.BackColor = txtNotHighlight.BackColor
        txtTargetTimeout.ForeColor = txtNotHighlight.ForeColor
    End If
    If optPurgeTarget.Value = cYES Then
        optPurgeTarget.BackColor = frmHighlight.BackColor
        optPurgeTarget.ForeColor = frmHighlight.ForeColor
        txtPurgeTarget.BackColor = frmHighlight.ForeColor
        txtPurgeTarget.ForeColor = frmHighlight.BackColor
        txtTargetTimeout.BackColor = frmHighlight.ForeColor
        txtTargetTimeout.ForeColor = frmHighlight.BackColor
    Else
        optPurgeTarget.BackColor = frmNotHighlight.BackColor
        optPurgeTarget.ForeColor = frmNotHighlight.ForeColor
        txtPurgeTarget.BackColor = txtNotHighlight.BackColor
        txtPurgeTarget.ForeColor = txtNotHighlight.ForeColor
        txtTargetTimeout.BackColor = txtNotHighlight.BackColor
        txtTargetTimeout.ForeColor = txtNotHighlight.ForeColor
    End If
    
    ' LOAD METHOD
    If optNoLoad.Value = cYES Then
        optNoLoad.BackColor = frmHighlight.BackColor
        optNoLoad.ForeColor = frmHighlight.ForeColor
    Else
        optNoLoad.BackColor = frmNotHighlight.BackColor
        optNoLoad.ForeColor = frmNotHighlight.ForeColor
    End If
    If optLoadTime.Value = cYES Then
        optLoadTime.BackColor = frmHighlight.BackColor
        optLoadTime.ForeColor = frmHighlight.ForeColor
        txtLoadTime.BackColor = frmHighlight.ForeColor
        txtLoadTime.ForeColor = frmHighlight.BackColor
    Else
        optLoadTime.BackColor = frmNotHighlight.BackColor
        optLoadTime.ForeColor = frmNotHighlight.ForeColor
        txtLoadTime.BackColor = txtNotHighlight.BackColor
        txtLoadTime.ForeColor = txtNotHighlight.ForeColor
    End If
    If optWcm.Value = cYES Then
        optWcm.BackColor = frmHighlight.BackColor
        optWcm.ForeColor = frmHighlight.ForeColor
        txtWorkCapMult.BackColor = frmHighlight.ForeColor
        txtWorkCapMult.ForeColor = frmHighlight.BackColor
        txtEPAFill.BackColor = frmHighlight.ForeColor
        txtEPAFill.ForeColor = frmHighlight.BackColor
    Else
        optWcm.BackColor = frmNotHighlight.BackColor
        optWcm.ForeColor = frmNotHighlight.ForeColor
        txtWorkCapMult.BackColor = txtNotHighlight.BackColor
        txtWorkCapMult.ForeColor = txtNotHighlight.ForeColor
        txtEPAFill.BackColor = txtNotHighlight.BackColor
        txtEPAFill.ForeColor = txtNotHighlight.ForeColor
    End If
    If optLoadweight.Value = cYES Then
        optLoadweight.BackColor = frmHighlight.BackColor
        optLoadweight.ForeColor = frmHighlight.ForeColor
        txtTargetWt.BackColor = frmHighlight.ForeColor
        txtTargetWt.ForeColor = frmHighlight.BackColor
    Else
        optLoadweight.BackColor = frmNotHighlight.BackColor
        optLoadweight.ForeColor = frmNotHighlight.ForeColor
        txtTargetWt.BackColor = txtNotHighlight.BackColor
        txtTargetWt.ForeColor = txtNotHighlight.ForeColor
    End If
    If optLoadBreakthrough.Value = cYES Then
        optLoadBreakthrough.BackColor = frmHighlight.BackColor
        optLoadBreakthrough.ForeColor = frmHighlight.ForeColor
        txtLoadBreakthrough.BackColor = frmHighlight.ForeColor
        txtLoadBreakthrough.ForeColor = frmHighlight.BackColor
    Else
        optLoadBreakthrough.BackColor = frmNotHighlight.BackColor
        optLoadBreakthrough.ForeColor = frmNotHighlight.ForeColor
        txtLoadBreakthrough.BackColor = txtNotHighlight.BackColor
        txtLoadBreakthrough.ForeColor = txtNotHighlight.ForeColor
    End If
    If optFIDBreakthrough.Value = cYES Then
        optFIDBreakthrough.BackColor = frmHighlight.BackColor
        optFIDBreakthrough.ForeColor = frmHighlight.ForeColor
        txtFIDmg.BackColor = frmHighlight.ForeColor
        txtFIDmg.ForeColor = frmHighlight.BackColor
    Else
        optFIDBreakthrough.BackColor = frmNotHighlight.BackColor
        optFIDBreakthrough.ForeColor = frmNotHighlight.ForeColor
        txtFIDmg.BackColor = txtNotHighlight.BackColor
        txtFIDmg.ForeColor = txtNotHighlight.ForeColor
    End If
    
    txtLoadRate.BackColor = txtNotHighlight.BackColor
    txtButnPercent.BackColor = txtNotHighlight.BackColor
    
    txtMaxLoadTime.BackColor = txtNotHighlight.BackColor
    
    chkOrvrMfc.BackColor = frmNotHighlight.BackColor
    
    ' END METHOD
    If optEndCycles.Value = cYES Then
        optEndCycles.BackColor = frmHighlight.BackColor
        optEndCycles.ForeColor = frmHighlight.ForeColor
        txtPFCycle.BackColor = frmHighlight.ForeColor
        txtPFCycle.ForeColor = frmHighlight.BackColor
    Else
        optEndCycles.BackColor = frmNotHighlight.BackColor
        optEndCycles.ForeColor = frmNotHighlight.ForeColor
        txtPFCycle.BackColor = txtNotHighlight.BackColor
        txtPFCycle.ForeColor = txtNotHighlight.ForeColor
    End If
    If optEndWeightChange.Value = cYES Then
        optEndWeightChange.BackColor = frmHighlight.BackColor
        optEndWeightChange.ForeColor = frmHighlight.ForeColor
        txtWeightChangeTol.BackColor = frmHighlight.BackColor
        txtWeightChangeTol.ForeColor = frmHighlight.ForeColor
        txtConsecutiveCycles.BackColor = frmHighlight.BackColor
        txtConsecutiveCycles.ForeColor = frmHighlight.ForeColor
        txtMaximumCycles.BackColor = frmHighlight.BackColor
        txtMaximumCycles.ForeColor = frmHighlight.ForeColor
        txtMinimumCycles.BackColor = frmHighlight.BackColor
        txtMinimumCycles.ForeColor = frmHighlight.ForeColor
    Else
        optEndWeightChange.BackColor = frmNotHighlight.BackColor
        optEndWeightChange.ForeColor = frmNotHighlight.ForeColor
        txtWeightChangeTol.BackColor = txtNotHighlight.BackColor
        txtWeightChangeTol.ForeColor = txtNotHighlight.ForeColor
        txtConsecutiveCycles.BackColor = txtNotHighlight.BackColor
        txtConsecutiveCycles.ForeColor = txtNotHighlight.ForeColor
        txtMaximumCycles.BackColor = txtNotHighlight.BackColor
        txtMaximumCycles.ForeColor = txtNotHighlight.ForeColor
        txtMinimumCycles.BackColor = txtNotHighlight.BackColor
        txtMinimumCycles.ForeColor = txtNotHighlight.ForeColor
    End If
    
    
    chkPurgeAuxCan.BackColor = frmNotHighlight.BackColor
    txtPurgeFlow.BackColor = txtNotHighlight.BackColor
    txtPurgeVolume.BackColor = txtNotHighlight.BackColor
    txtPrimaryScaleNo.BackColor = txtNotHighlight.BackColor
    txtAuxScaleNo.BackColor = txtNotHighlight.BackColor
    txtPauseLeakTime.BackColor = txtNotHighlight.BackColor
    txtPauseLoadTime.BackColor = txtNotHighlight.BackColor
    txtPausePurgeTime.BackColor = txtNotHighlight.BackColor
    
    ' purge-to-target options
    If optTargetContinuous.Value Then
        optTargetContinuous.BackColor = frmHighlight.BackColor
        optTargetContinuous.ForeColor = frmHighlight.ForeColor
    Else
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
    End If
    If optTargetPurgePauseRepeat.Value Then
        optTargetPurgePauseRepeat.BackColor = frmHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmHighlight.ForeColor
        txtTargetPurge.BackColor = frmHighlight.ForeColor
        txtTargetPurge.ForeColor = frmHighlight.BackColor
        txtTargetPause.BackColor = frmHighlight.ForeColor
        txtTargetPause.ForeColor = frmHighlight.BackColor
    Else
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
    End If
    
'    frmAnalyzer.txtTargetConcentration.BackColor = txtNotHighlight.BackColor
'    frmAnalyzer.txtDwellTime.BackColor = txtNotHighlight.BackColor
    
    txtIDLoad.BackColor = txtNotHighlight.BackColor
    txtIDPurge.BackColor = txtNotHighlight.BackColor
    txtIDVent.BackColor = txtNotHighlight.BackColor
    txtLoadL.BackColor = txtNotHighlight.BackColor
    txtPurgeL.BackColor = txtNotHighlight.BackColor
    txtVentL.BackColor = txtNotHighlight.BackColor
    
    txtNitrogenFlow.BackColor = txtNotHighlight.BackColor
    txtLiveFuelChgFreq.BackColor = txtNotHighlight.BackColor
    chkADF_Heater.BackColor = frmNotHighlight.BackColor
    txtADF_HeaterSP.BackColor = txtNotHighlight.BackColor
    
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

Private Function ValidRecipe() As Boolean
' Function Name:    ValidRecipe
' Description:      Checks the validity of recipe settings.
'                   Used before saving the recipe file.
'                   Returns a true value if values are okay.
'                   Returns a false value if values are not okay.
'                   If an error is detected, an appropriate message
'                   is displayed.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 1
Dim cycleText As String
Dim Message As String
Dim flag As Boolean
Dim inc As Integer
Dim minScaleNum, maxScaleNum As Single
Dim dDate As Date
Dim testout, Target, temp1, temp2 As Single
Dim sGramsPerLiter As Single
Dim maxmass As Single
Dim maxrate As Single
Dim minrate As Single
Dim maxtime As Single


ValidRecipe = True
lblMessage.Caption = ""
minScaleNum = IIf(RecipeMode = STATIONMODE, 1, 0)
maxScaleNum = CInt(NR_SCALES)

' Name
If Len(txtRecipeName.text) > 50 Then
    ValidRecipe = False
    txtRecipeName.BackColor = EntryInvalid_BackColor
    Message = "Name is too long. 50 char Max"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
ElseIf Len(txtRecipeName.text) < 1 Then
    txtRecipeName.text = " "
End If

' Start Method
inc = 0
If optStartNow.Value = cYES Then inc = inc + 1
If optStartAfter.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtStartAfterMin, 1, 999, "Delay Start Minutes") Then ValidRecipe = False
Else
    If Not IsNumeric(txtStartAfterMin.text) Then txtStartAfterMin.text = "0"
    If Not Range_Check(txtStartAfterMin, 0, 999, "Delay Start Minutes") Then ValidRecipe = False
End If
If optStartAt.Value = cYES Then
    inc = inc + 1
    dDate = Now() + TimeSerial(0, 1, 0)
    If Not IsDate(txtStartAtDate.text) Then
        ValidRecipe = False
        txtStartAtDate.BackColor = EntryInvalid_BackColor
        Message = "Invalid Start At Date"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    ElseIf CDate(txtStartAtDate.text) < dDate Then
        ValidRecipe = False
        txtStartAtDate.BackColor = EntryInvalid_BackColor
        Message = "Start At Date must be later than Now"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    End If
Else
    If Not IsDate(txtStartAtDate.text) Then txtStartAtDate.text = Format(dDate, "M/D/YYYY hh:mm")
    If CDate(txtStartAtDate.text) < dDate Then txtStartAtDate.text = Format(dDate, "M/D/YYYY hh:mm")
End If
If inc = 0 Then
    ValidRecipe = False
    optStartNow.BackColor = EntryInvalid_BackColor
    Message = "Must Check ONE Start Method Box"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
ElseIf inc > 1 Then
    ValidRecipe = False
    If optStartNow.Value > 0 Then optStartNow.BackColor = EntryInvalid_BackColor
    If optStartAfter.Value > 0 Then optStartAfter.BackColor = EntryInvalid_BackColor
    If optStartAt.Value > 0 Then optStartAt.BackColor = EntryInvalid_BackColor
    Message = "Only ONE Start Method may be Selected at a Time"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
End If

' End Method
inc = 0
If (tabsCycletype.Tabs(CyclePurgeLoad).Selected) Then cycleText = CycleTypeDesc(CyclePurgeLoad)
If optEndCycles.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtPFCycle, 1, 999, cycleText) Then ValidRecipe = False
Else
    If Not IsNumeric(txtPFCycle.text) Then txtPFCycle.text = "0"
    If Not Range_Check(txtPFCycle, 0, 999, cycleText) Then ValidRecipe = False
    If (StationCanister(DispStn, DispShift).WorkingCapacity = 0) Then
        ValidRecipe = False
        optEndCycles.BackColor = EntryInvalid_BackColor
        Message = "Canister must have a Working Capacity for this End Method"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    End If
End If
If optEndWeightChange = cYES Then
    inc = inc + 1
    If Not Range_Check(txtWeightChangeTol, 0.1, 100, "Wt Change Tolerance") Then ValidRecipe = False
    If Not Range_Check(txtConsecutiveCycles, 1, 99, "Consecutive Cycles") Then ValidRecipe = False
    If Not Range_Check(txtMinimumCycles, 1, 99, "Minimum Cycles") Then ValidRecipe = False
    If Not Range_Check(txtMaximumCycles, 1, 99, "Maximum Cycles") Then ValidRecipe = False
    If ValidRecipe Then
        If CInt(txtMinimumCycles.text) > CInt(txtMaximumCycles.text) Then
            ValidRecipe = False
            txtMinimumCycles.BackColor = EntryInvalid_BackColor
            txtMaximumCycles.BackColor = EntryInvalid_BackColor
            Message = "Minimum must be less than Max"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        End If
    End If
    If ValidRecipe Then
        If CInt(txtConsecutiveCycles.text) > CInt(txtMaximumCycles.text) Then
            ValidRecipe = False
            txtConsecutiveCycles.BackColor = EntryInvalid_BackColor
            txtMaximumCycles.BackColor = EntryInvalid_BackColor
            Message = "Consecutive must be less than Max"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        End If
    End If
Else
    If Not IsNumeric(txtWeightChangeTol.text) Then txtWeightChangeTol.text = "0"
    If Not IsNumeric(txtConsecutiveCycles.text) Then txtConsecutiveCycles.text = "0"
    If Not IsNumeric(txtMinimumCycles.text) Then txtMinimumCycles.text = "0"
    If Not IsNumeric(txtMaximumCycles.text) Then txtMaximumCycles.text = "0"
End If
If inc = 0 Then
    ValidRecipe = False
    optEndCycles.BackColor = EntryInvalid_BackColor
    Message = "Must Check ONE End Method Box"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
ElseIf inc > 1 Then
    ValidRecipe = False
    If optEndCycles.Value > 0 Then optEndCycles.BackColor = EntryInvalid_BackColor
    If optEndWeightChange.Value > 0 Then optEndWeightChange.BackColor = EntryInvalid_BackColor
    Message = "Only ONE End Method may be Selected at a Time"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
End If

' Live Fuel Options
If systemhasLIVEFUEL And chkLiveFuel.Value = cYES Then
    ' Live Fuel Recipe
    If chkADF_Heater.Value = cYES Then
        If (chkADF_Heater.Caption = "Use WaterBath") Then
            If USINGC Then
                 If Not Range_Check(txtADF_HeaterSP, WB_AIO.EuMin, WB_AIO.EuMax, "LiveFuel WaterBath SetPoint") Then ValidRecipe = False
            End If
            If USINGF Then
                 If Not Range_Check(txtADF_HeaterSP, DegCtoF(WB_AIO.EuMin), DegCtoF(WB_AIO.EuMax), "LiveFuel WaterBath SetPoint") Then ValidRecipe = False
            End If
        Else
            If USINGC Then
                 If Not Range_Check(txtADF_HeaterSP, 15, 50, "LiveFuel Heater SetPoint") Then ValidRecipe = False
            End If
            If USINGF Then
                 If Not Range_Check(txtADF_HeaterSP, 60, 120, "LiveFuel Heater SetPoint") Then ValidRecipe = False
            End If
        End If
    End If
    If Not Range_Check(txtLiveFuelChgFreq, 1, 999, "LiveFuel Change Frequency") Then ValidRecipe = False
    ' Auto Drain/Fill Options
    If chkLiveFuelChgAuto.Value = cYES Then
        Select Case chkLoadRatePID.Value
            Case cYES
                ' can be zero since Load Rate PID is being used
                If Not Range_Check(txtNitrogenFlow, CSng(0), 50, "LiveFuel Vapor Flow Rate") Then ValidRecipe = False
            Case cNO
                ' can not be zero
                If Not Range_Check(txtNitrogenFlow, CSng(0.1), 50, "LiveFuel Vapor Flow Rate") Then ValidRecipe = False
        End Select
    End If
    If Not Range_Check(txtLoadRate, 0, 5000, "Target Load Rate") Then ValidRecipe = False
    If Not IsNumeric(txtButnPercent.text) Then txtButnPercent.text = "0"
Else
    ' Not a Live Fuel Recipe
    If Not IsNumeric(txtLiveFuelChgFreq.text) Then txtLiveFuelChgFreq.text = "0"
    If Not IsNumeric(txtNitrogenFlow.text) Then txtNitrogenFlow.text = "0"
    If Not Range_Check(txtLoadRate, 0, 10000, "Target Load Rate") Then ValidRecipe = False
    If Not Range_Check(txtButnPercent, 0, 100, "Percent Butane") Then ValidRecipe = False
End If

' Butane Load Options
If USINGLOADTIMELIMIT Then
    If optNoLoad.Value = 0 Then
        If Not Range_Check(txtMaxLoadTime, 0, 9999, "Max Load Time") Then ValidRecipe = False
    End If
End If

' Purge Values
If optPurgeTime.Value = cYES Then
    If Not Range_Check(txtPurgeFlow, 0.1, 100, "Purge Flow Rate") Then ValidRecipe = False
    If Not Range_Check(txtPurgeTime, 1, 9999, "Purge Time") Then ValidRecipe = False
'    If Not Range_Check(txtPurgeVolume, 0, 999, "Canister Volumes") Then ValidRecipe = False
End If
If optPurgeAuxOnly.Value = cYES Then
'    If Not Range_Check(txtPurgeFlow, 0, 100, "Purge Flow Rate") Then ValidRecipe = False
    If Not Range_Check(txtPurgeAuxOnly, 0.1, 9999, "Purge Aux Only Time") Then ValidRecipe = False
    If (chkPurgeAuxCan.Value = cNO) Then
        ValidRecipe = False
        chkPurgeAuxCan.BackColor = EntryInvalid_BackColor
        Message = "Must Select Purge Aux Canister"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    End If
    If (chkPurgeCansInSeries.Value = cYES) Then
        ValidRecipe = False
        chkPurgeCansInSeries.BackColor = EntryInvalid_BackColor
        Message = "No Series Purge with Aux Only"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    End If
End If
If optPurgeVolume.Value = cYES Then
    If Not Range_Check(txtPurgeFlow, 0.01, 100, "Purge Flow Rate") Then ValidRecipe = False
'    If Not Range_Check(txtPurgeTime, 0, 9999, "Purge Time") Then ValidRecipe = False
    If Not Range_Check(txtPurgeVolume, 0.01, 9999, "Canister Volumes") Then ValidRecipe = False
    If optTargetPurgePauseRepeat.Value Then
        If Not Range_Check(txtTargetPurge, 0.1, 999.9, "Purging Duration") Then ValidRecipe = False
        If Not Range_Check(txtTargetPause, 0.1, 999.9, "Pauseing Duration") Then ValidRecipe = False
    End If
End If
If optPurgeWC.Value = cYES Then
    If Not Range_Check(txtPurgeWC, 1, 110, "Purge WC Target") Then ValidRecipe = False
    If Not Range_Check(txtTargetTimeout, 1, 999, "Purge Timeout") Then ValidRecipe = False
    If optTargetPurgePauseRepeat.Value Then
        If Not Range_Check(txtTargetPurge, 0.1, 999.9, "Purging Duration") Then ValidRecipe = False
        If Not Range_Check(txtTargetPause, 0.1, 999.9, "Pauseing Duration") Then ValidRecipe = False
    End If
End If
If optPurgeTarget.Value = cYES Then
    If Not Range_Check(txtPurgeTarget, -19999, 19999, "Purge Target Weight") Then ValidRecipe = False
    If Not Range_Check(txtTargetTimeout, 1, 999, "Purge Timeout") Then ValidRecipe = False
    If optTargetPurgePauseRepeat.Value Then
        If Not Range_Check(txtTargetPurge, 0.1, 999.9, "Purging Duration") Then ValidRecipe = False
        If Not Range_Check(txtTargetPause, 0.1, 999.9, "Pauseing Duration") Then ValidRecipe = False
    End If
End If
If optPurgeUndo.Value = cYES Then
    If Not Range_Check(txtTargetTimeout, 1, 999, "Purge Timeout") Then ValidRecipe = False
    If optTargetPurgePauseRepeat.Value Then
        If Not Range_Check(txtTargetPurge, 0.1, 999.9, "Purging Duration") Then ValidRecipe = False
        If Not Range_Check(txtTargetPause, 0.1, 999.9, "Pauseing Duration") Then ValidRecipe = False
    End If
End If
If ((chkPurgeCansInSeries.Value = cON) And (chkPurgeAuxCan.Value = cOFF)) Then
    ValidRecipe = False
    chkPurgeAuxCan.BackColor = EntryInvalid_BackColor
    Message = "No Series Purge without Aux"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
End If
If chkUsePurgeOven.Value = cON Then
    If Not Range_Check(txtPurgeOvenSP, 0, 60, "Purge Oven SP") Then ValidRecipe = False
End If


' LeakCheck
If chkLeakCheck.Value = cYES Then
    ' Scales
    If USINGAUXLEAKCHECK Then
        If ((chkLeakPrimary.Value = cNO) And (chkLeakAux.Value = cNO)) Then
            ValidRecipe = False
            chkLeakPrimary.BackColor = EntryInvalid_BackColor
            chkLeakAux.BackColor = EntryInvalid_BackColor
            Message = "Select Pri OR Aux OR Both"
            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        Else
            If ((chkLeakPrimary.Value = cYES) And (chkPrimaryScale.Value = cNO)) Then
                    ValidRecipe = False
                    chkLeakPrimary.BackColor = EntryInvalid_BackColor
                    chkPrimaryScale.BackColor = EntryInvalid_BackColor
                    Message = "Primary Scale required for LeakCheck of Primary Scale"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
            If ((chkLeakAux.Value = cYES) And (chkUseAuxScale.Value = cNO)) Then
                    ValidRecipe = False
                    chkLeakAux.BackColor = EntryInvalid_BackColor
                    chkPrimaryScale.BackColor = EntryInvalid_BackColor
                    Message = "Aux Scale required for LeakCheck of Aux Scale"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
    End If
End If

' Scale Numbers
If chkUseAuxScale.Value = cYES Then
    If Not Range_Check(txtAuxScaleNo, minScaleNum, maxScaleNum, "Aux Scale Number") Then ValidRecipe = False
Else
    If Not IsNumeric(txtAuxScaleNo.text) Then txtAuxScaleNo.text = "0"
    If CInt(txtAuxScaleNo.text) > maxScaleNum Then txtAuxScaleNo.text = "0"
End If
If chkPrimaryScale.Value = cYES Then
    If Not Range_Check(txtPrimaryScaleNo, minScaleNum, maxScaleNum, "Primary Scale Number") Then ValidRecipe = False
Else
    If Not IsNumeric(txtPrimaryScaleNo.text) Then txtPrimaryScaleNo.text = "0"
    If CInt(txtPrimaryScaleNo.text) > maxScaleNum Then txtPrimaryScaleNo.text = "0"
End If

' Pause Values
If chkPauseAfterLeak.Value = cNO And Not IsNumeric(txtPauseLeakTime.text) Then txtPauseLeakTime.text = "0"
If chkPauseAfterLeak.Value = cYES Then
    If Not Range_Check(txtPauseLeakTime, 0.1, 9999, "Pause After Leak") Then ValidRecipe = False
End If
If (Not optPauseAfterLoad.Value) And Not IsNumeric(txtPauseLoadTime.text) Then txtPauseLoadTime.text = "0"
If (optPauseAfterLoad.Value) Then
    If Not Range_Check(txtPauseLoadTime, 0.1, 9999, "Pause After Load") Then ValidRecipe = False
End If
If (Not optPauseAfterPurge.Value) And Not IsNumeric(txtPausePurgeTime.text) Then txtPausePurgeTime.text = "0"
If (optPauseAfterPurge.Value) Then
    If Not Range_Check(txtPausePurgeTime, 0.1, 9999, "Pause After Purge") Then ValidRecipe = False
End If

' Load Method options
inc = 0
If optNoLoad.Value = cYES Then inc = inc + 1
If optLoadTime.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtLoadTime, 1, 9999, "Load by Time Minutes") Then ValidRecipe = False
Else
    If Not IsNumeric(txtLoadTime.text) Then txtLoadTime.text = "0"
    If Not Range_Check(txtLoadTime, 0, 9999, "Load by Time Minutes") Then ValidRecipe = False
End If
If optWcm.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtWorkCapMult, 0.1, 99.9, "WC Multiplier") Then ValidRecipe = False
    If Not Range_Check(txtEPAFill, 0, 36, "WC Hours") Then ValidRecipe = False
Else
    If Not IsNumeric(txtWorkCapMult.text) Then txtWorkCapMult.text = "0"
    If Not Range_Check(txtWorkCapMult, 0, 99.9, "WC Multiplier") Then ValidRecipe = False
    If Not IsNumeric(txtEPAFill.text) Then txtEPAFill.text = "0"
    If Not Range_Check(txtEPAFill, 0, 36, "WC Hours") Then ValidRecipe = False
End If
If optLoadweight.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtTargetWt, 0.01, 1500, "Load by Weight Grams") Then ValidRecipe = False
Else
    If Not IsNumeric(txtTargetWt.text) Then txtTargetWt.text = "0"
    If Not Range_Check(txtTargetWt, 0, 1500, "Load by Weight Grams") Then ValidRecipe = False
End If
If optLoadBreakthrough.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtLoadBreakthrough, 0.01, 1500, "Load by Breakthrough") Then ValidRecipe = False
Else
    If Not IsNumeric(txtLoadBreakthrough.text) Then txtLoadBreakthrough.text = "0"
    If Not Range_Check(txtLoadBreakthrough, 0, 1500, "Load by Breakthrough") Then ValidRecipe = False
End If
If optFIDBreakthrough.Value = cYES Then
    inc = inc + 1
    If Not Range_Check(txtFIDmg, 0.01, 99.9, "FID Breakthrough") Then ValidRecipe = False
'    If chkUseAnalyzer.Value = cNO Then
'        ValidRecipe = False
'        chkUseAnalyzer.BackColor = EntryInvalid_BackColor
'        Message = "Must Check Fid Analyzer for FID Breakthrough"
'        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
'    End If
Else
    If Not IsNumeric(txtFIDmg.text) Then txtFIDmg.text = "0"
    If Not Range_Check(txtFIDmg, 0, 99.9, "FID Breakthrough") Then ValidRecipe = False
End If
If inc = 0 Then
    ValidRecipe = False
    optNoLoad.BackColor = EntryInvalid_BackColor
    Message = "Must Check ONE Load Box"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
ElseIf inc > 1 Then
    ValidRecipe = False
    If optNoLoad.Value = cYES Then optNoLoad.BackColor = EntryInvalid_BackColor
    If optLoadTime.Value = cYES Then optLoadTime.BackColor = EntryInvalid_BackColor
    If optWcm.Value = cYES Then optWcm.BackColor = EntryInvalid_BackColor
    If optLoadweight.Value = cYES Then optLoadweight.BackColor = EntryInvalid_BackColor
    If optLoadBreakthrough.Value = cYES Then optLoadBreakthrough.BackColor = EntryInvalid_BackColor
    If optFIDBreakthrough.Value = cYES Then optFIDBreakthrough.BackColor = EntryInvalid_BackColor
    Message = "Only ONE Load Option may be Selected at a Time"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
End If


' Line Volume Values
If USINGLINEVOLUME Then
    flag = True
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
    Else
        ValidRecipe = False
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


' ***********************************************************************
' Additional Validation Checks when saving a recipe to a specific station
' ***********************************************************************
If RecipeMode = STATIONMODE Then

    If ValidRecipe Then
        If chkLiveFuel.Value = cYES Then
            If ((STN_INFO(DispStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE)) Then
                ' Use Live Fuel MFC
                If chkOrvrMfc.Value = cYES Then
                    If (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE) Then
                        ' Use High Range Live Fuel MFC
                        Select Case chkLoadRatePID.Value
                            Case cYES
                                ' can be zero since Load Rate PID is being used
                                If Not Range_Check(txtNitrogenFlow, CSng(0), (0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "Fuel Vapor Flow Rate") Then ValidRecipe = False
                            Case cNO
                                ' can not be zero
                                If Not Range_Check(txtNitrogenFlow, (MfcSpMin * 0.01 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), (0.95 * Stn_AIO(DispStn, asLiveFuelVaporORVRFlowSP).EuMax), "Fuel Vapor Flow Rate") Then ValidRecipe = False
                        End Select
                    Else
                        ValidRecipe = False
                        chkOrvrMfc.BackColor = EntryInvalid_BackColor
                        Message = "This Station doesn't have Dual Range MFC's"
                        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                    End If
                Else
                    ' use Low Range Live Fuel MFC
                    Select Case chkLoadRatePID.Value
                        Case cYES
                            ' can be zero since Load Rate PID is being used
                            If Not Range_Check(txtNitrogenFlow, CSng(0), (0.95 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "Fuel Vapor Flow Rate") Then ValidRecipe = False
                        Case cNO
                            ' can not be zero
                            If Not Range_Check(txtNitrogenFlow, (MfcSpMin * 0.01 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), (0.95 * Stn_AIO(DispStn, asLiveFuelVaporFlowSP).EuMax), "Fuel Vapor Flow Rate") Then ValidRecipe = False
                    End Select
                End If
            Else
                ValidRecipe = False
                chkLiveFuel.BackColor = EntryInvalid_BackColor
                Message = "This Station doesn't support Live Fuel"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        Else
            ' Nitrogen MFC
            ' can be zero since butane is being used
            If chkOrvrMfc.Value = cYES Then
                If ((STN_INFO(DispStn).Type = STN_ORVR2_TYPE) Or (STN_INFO(DispStn).Type = STN_LIVEORVR2_TYPE)) Then
                    ' Use High Range Nitrogen MFC
                    ' can be zero since butane is being used
                    If Not Range_Check(txtNitrogenFlow, CSng(0), (0.95 * Stn_AIO(DispStn, asNitrogenORVRFlowSP).EuMax), "Nitrogen Flow Rate") Then ValidRecipe = False
                Else
                    ValidRecipe = False
                    chkOrvrMfc.BackColor = EntryInvalid_BackColor
                    Message = "This Station doesn't have Dual Range MFC's"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                End If
            Else
                ' use Low Range Nitrogen MFC
                If Not Range_Check(txtNitrogenFlow, CSng(0), (0.95 * Stn_AIO(DispStn, asNitrogenFlowSP).EuMax), "Nitrogen Flow Rate") Then ValidRecipe = False
            End If
        End If
    End If
    
    ' Live Fuel LoadRate PID
    If ValidRecipe Then
        If chkLoadRatePID.Value = cYES Then
            If ((STN_INFO(DispStn).Type <> STN_LIVEFUEL_TYPE) And (STN_INFO(DispStn).Type <> STN_LIVEREG_TYPE) And (STN_INFO(DispStn).Type <> STN_LIVEORVR2_TYPE)) Then
                ValidRecipe = False
                chkLoadRatePID.BackColor = EntryInvalid_BackColor
                Message = "Load Rate PID is only available for Live Fuel"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
    End If

    ' Scale Numbers
    If ValidRecipe Then
        If USINGHARDPIPEDSCALES Then
            If chkUseAuxScale.Value = cNO Then
                chkUseAuxScale.BackColor = EntryInvalid_BackColor
                Message = "Aux Scale must be Enabled"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                ValidRecipe = False
            ElseIf (CInt(txtAuxScaleNo.text) <> STN_INFO(DispStn).DefAuxScale) Then
                    ValidRecipe = False
                    txtAuxScaleNo.BackColor = EntryInvalid_BackColor
                    Message = "Aux Scale must be Scale #" & Format(STN_INFO(DispStn).DefAuxScale, "#0")
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
            If chkPrimaryScale.Value = cNO Then
                chkPrimaryScale.BackColor = EntryInvalid_BackColor
                Message = "Aux Scale must be Enabled"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                ValidRecipe = False
            ElseIf (CInt(txtPrimaryScaleNo.text) <> STN_INFO(DispStn).DefPriScale) Then
                    ValidRecipe = False
                    txtPrimaryScaleNo.BackColor = EntryInvalid_BackColor
                    Message = "Primary Scale must be Scale #" & Format(STN_INFO(DispStn).DefPriScale, "#0")
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        Else
            If chkUseAuxScale.Value = cYES Then
                If CInt(txtAuxScaleNo.text) >= FIRST_REMOTESCALE Then
                    Message = "Aux Scale can't be a Remote Scale"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                    ValidRecipe = False
                End If
            End If
            If chkUseAuxScale.Value = cYES And chkPrimaryScale.Value = cYES Then
                If CInt(txtAuxScaleNo.text) = CInt(txtPrimaryScaleNo.text) Then
                    ValidRecipe = False
                    txtAuxScaleNo.BackColor = EntryInvalid_BackColor
                    txtPrimaryScaleNo.BackColor = EntryInvalid_BackColor
                    Message = "Primary & Aux Scales can't be the same scale"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                End If
            End If
        End If
    End If
    
    If ValidRecipe Then
        If optLoadweight.Value = cYES Then
            If chkPrimaryScale.Value = cNO Then
                ValidRecipe = False
                chkPrimaryScale.BackColor = EntryInvalid_BackColor
                Message = "Primary Scale required for Load By Weight"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            ElseIf (ValueFromText(txtTargetWt.text) > (1.2 * StationCanister(DispStn, DispShift).WorkingCapacity)) Then
                ValidRecipe = False
                txtTargetWt.BackColor = EntryInvalid_BackColor
                Message = "Load Weight exceeds Canister Capacity"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        If optLoadBreakthrough.Value = cYES Then
            If chkUseAuxScale.Value = cNO Then
                ValidRecipe = False
                chkUseAuxScale.BackColor = EntryInvalid_BackColor
                Message = "Aux Scale required for Load By Breakthrough"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        If chkPurgeAuxCan.Value = cYES Then
            If chkUseAuxScale.Value = cNO Then
                ValidRecipe = False
                chkUseAuxScale.BackColor = EntryInvalid_BackColor
                Message = "Must have AUX Canister Checked"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        If optWcm.Value = cYES And STN_INFO(DispStn).Type <> STN_LIVEFUEL_TYPE Then
            ' * using WC multiplier set it up *
            Target = (StationCanister(DispStn, DispShift).WorkingCapacity * CSng(txtWorkCapMult.text))
            If CSng(txtEPAFill.text) < ((Target / CSng(txtLoadRate.text)) - 0.1) Then
                ' Not enough time for the specified Load Rate
                ' Recalc flow to achieve time line
                If STN_INFO(DispStn).Type = STN_ORVR2_TYPE And chkOrvrMfc Then
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                    temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
                Else
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                    temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                End If
                temp2 = (Target / CSng(txtEPAFill.text))
                If temp2 > temp1 Then
                    ' MFC can't flow enough to meet time limits; set Load Rate to Max
                    ValidRecipe = False
                    txtLoadRate.text = CStr(temp1)
                    txtLoadRate.BackColor = EntryInvalid_BackColor
                    txtEPAFill.BackColor = EntryInvalid_BackColor
                    txtWorkCapMult.BackColor = EntryInvalid_BackColor
                    Message = "Adjusted LOAD FLOW RATE to Max" & vbCrLf & "It is still too low to meet the time limits!"
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                Else
                    ' MFC can meet time limits; adjust Load Rate accordingly
                    txtLoadRate.text = Format(temp2, "##0.00")
                    Message = "Adjusted LOAD FLOW RATE to " & txtLoadRate.text
                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                End If
            End If
        End If
        
        ' Purge Rate
        If optNoPurge.Value = cNO And optPurgeAuxOnly.Value = cNO And optPurgeProfile.Value = cNO And ValidRecipe Then
            temp1 = Stn_AIO(DispStn, asPurgeAirFlow).EuMax * 0.01 * MfcSpMin
            temp2 = Stn_AIO(DispStn, asPurgeAirFlow).EuMax * 0.95
            If Not Range_Check(txtPurgeFlow, temp1, temp2, "Purge Flow Rate") Then ValidRecipe = False
        End If
        
        ' Purge Timeout (Max Volumes)
        If optPurgeVolume.Value = cYES And ValidRecipe Then
            temp1 = ValueFromText(txtTargetTimeout.text)
            temp2 = ValueFromText(txtPurgeVolume.text)
            If (temp1 < (1.1 * temp2)) Then
                txtTargetTimeout.text = Format((1.1 * temp2), "####0")
                Message = "Adjusted TARGET TIMEOUT to " & txtTargetTimeout.text
                If Not NotDebugMMW Then lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        
        ' Purge Scales
        If ((optPurgeWC.Value = cYES Or optPurgeTarget.Value = cYES Or optPurgeUndo.Value = cYES) And ValidRecipe) Then
            If (chkPrimaryScale.Value = cNO) Then
                ValidRecipe = False
                chkPrimaryScale.BackColor = EntryInvalid_BackColor
                Message = "Primary Scale required for "
                If (optPurgeWC.Value = cYES) Then Message = Message & PurgeMethodDesc(PURGEBYWC)
                If (optPurgeTarget.Value = cYES) Then Message = Message & PurgeMethodDesc(PURGETOTARGET)
                If (optPurgeUndo.Value = cYES) Then Message = Message & PurgeMethodDesc(PURGETOUNDOLOAD)
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        
        ' Load Rate
        If optNoLoad.Value = cNO And ValidRecipe Then
            Select Case STN_INFO(DispStn).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE
                    sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                    ' Butane is mixed with the N2
                    temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                    temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                    If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                    If Not Range_Check(txtButnPercent, 0.1, 100, "Percent Butane") Then ValidRecipe = False
                    ' now is the mix % greater than the capabilities of the flow controller
                    If ValidRecipe Then
                        temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                        temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                        If temp1 > temp2 Then
                            ValidRecipe = False
                            txtLoadRate.BackColor = EntryInvalid_BackColor
                            txtButnPercent.BackColor = EntryInvalid_BackColor
                            Message = "Required N2 flow(" & Format(temp1, "####0.000") & ") exceeds MFC range(" & Format(temp2, "###0.000") & ")."
                            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                        End If
                    End If
                Case STN_ORVR2_TYPE
                    ' Butane is mixed with the N2
                    If chkOrvrMfc.Value = cYES Then
                        ' use higher range MFC
                        sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                        temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                        temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
                    Else
                        ' use lower range MFC
                        sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                        temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                        temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                    End If
                    If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                    If Not Range_Check(txtButnPercent, 0.1, 100, "Percent Butane") Then ValidRecipe = False
                    ' now is the mix % greater than the capabilities of the flow controller
                    If ValidRecipe Then
                        If chkOrvrMfc.Value = cYES Then
                            ' use higher range MFC
                            sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfc2DensityMult
                            temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                            temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenORVRFlow).EuMax
                        Else
                            ' use lower range MFC
                            sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                            temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                            temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                        End If
                        If temp1 > temp2 Then
                            ValidRecipe = False
                            txtLoadRate.BackColor = EntryInvalid_BackColor
                            txtButnPercent.BackColor = EntryInvalid_BackColor
                            Message = "Required N2 flow(" & Format(temp1, "####0.000") & ") exceeds MFC range(" & Format(temp2, "###0.000") & ")."
                            lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                        End If
                    End If
                Case STN_LIVEFUEL_TYPE
                    ' use Live Fuel
                    sGramsPerLiter = LiveFuelVaporDensity
                    temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, DeadLiveFuelDensity)) * 0.01 * MfcSpMin)
                    sGramsPerLiter = LiveFuelVaporDensity
                    temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, sGramsPerLiter)) * 0.95)
                    ' LiveFuel Vapor is carried by the Nitrogen
                    If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                    ' heater or waterbath ??
                    If ((Not STN_INFO(DispStn).ADF_DEF.hasADF_Heater) And (Not STN_INFO(DispStn).ADF_TANKTYPE = 90)) Then
                        chkADF_Heater.Value = cNO
                    End If
                Case STN_LIVEREG_TYPE
                    ' LiveFuel OR Butane/Nitrogen ??
                    If chkLiveFuel.Value = cYES Then
                        ' use Live Fuel
                        sGramsPerLiter = LiveFuelVaporDensity
                        temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                        temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, sGramsPerLiter)) * 0.95)
                        ' LiveFuel Vapor is carried by the Nitrogen
                        If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                        ' heater or waterbath ??
                        If ((Not STN_INFO(DispStn).ADF_DEF.hasADF_Heater) And (Not STN_INFO(DispStn).ADF_TANKTYPE = 90)) Then
                            chkADF_Heater.Value = cNO
                        End If
                    Else
                        ' use Butane/Nitrogen
                        sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                        ' Butane is mixed with the N2
                        temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                        temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                        If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                        If Not Range_Check(txtButnPercent, 0.1, 100, "Percent Butane") Then ValidRecipe = False
                        ' now is the mix % greater than the capabilities of the flow controller
                        If ValidRecipe Then
                            temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                            temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                            If temp1 > temp2 Then
                                ValidRecipe = False
                                txtLoadRate.BackColor = EntryInvalid_BackColor
                                txtButnPercent.BackColor = EntryInvalid_BackColor
                                Message = "Required N2 flow(" & Format(temp1, "####0.000") & ") exceeds MFC range(" & Format(temp2, "###0.000") & ")."
                                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                            End If
                        End If
                    End If
                Case STN_LIVEORVR2_TYPE
                    ' LiveFuel OR Butane/Nitrogen OR ORVR Butane/Nitrogen??
                    If chkLiveFuel.Value = cYES Then
                        ' use Live Fuel
                        sGramsPerLiter = LiveFuelVaporDensity
                        If chkOrvrMfc.Value = cYES Then
                            ' use ORVR LiveFuel/Nitrogen
                            temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporORVRFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                            temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
                            ' LiveFuel Vapor is carried by the Nitrogen
                            If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                            ' heater or waterbath ??
                            If ((Not STN_INFO(DispStn).ADF_DEF.hasADF_Heater) And (Not STN_INFO(DispStn).ADF_TANKTYPE = 90)) Then
                                chkADF_Heater.Value = cNO
                            End If
                        Else
                            ' use low range LiveFuel/Nitrogen
                            temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                            temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asLiveFuelVaporFlow).EuMax, sGramsPerLiter)) * 0.95)
                            ' LiveFuel Vapor is carried by the Nitrogen
                            If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                            ' heater or waterbath ??
                            If ((Not STN_INFO(DispStn).ADF_DEF.hasADF_Heater) And (Not STN_INFO(DispStn).ADF_TANKTYPE = 90)) Then
                                chkADF_Heater.Value = cNO
                            End If
                        End If
                    Else
                        ' use Butane/Nitrogen
                        If chkOrvrMfc.Value = cYES Then
                            ' use ORVR Butane/Nitrogen
                            sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                            ' Butane is mixed with the N2
                            temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                            temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneORVRFlow).EuMax, sGramsPerLiter)) * 0.95)
                            If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                            If Not Range_Check(txtButnPercent, 0.1, 100, "Percent Butane") Then ValidRecipe = False
                            ' now is the mix % greater than the capabilities of the flow controller
                            If ValidRecipe Then
                                temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                                temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenORVRFlow).EuMax
                                If temp1 > temp2 Then
                                    ValidRecipe = False
                                    txtLoadRate.BackColor = EntryInvalid_BackColor
                                    txtButnPercent.BackColor = EntryInvalid_BackColor
                                    Message = "Required N2 flow(" & Format(temp1, "####0.000") & ") exceeds MFC range(" & Format(temp2, "###0.000") & ")."
                                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                                End If
                            End If
                        Else
                            ' use low flow Butane/Nitrogen
                            sGramsPerLiter = GramsPerLiter * STN_INFO(DispStn).ButMfcDensityMult
                            ' Butane is mixed with the N2
                            temp1 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.01 * MfcSpMin)
                            temp2 = ((SlpmToGramsPerHour(Stn_AIO(DispStn, asButaneFlow).EuMax, sGramsPerLiter)) * 0.95)
                            If Not Range_Check(txtLoadRate, temp1, temp2, "Target Load Rate") Then ValidRecipe = False
                            If Not Range_Check(txtButnPercent, 0.1, 100, "Percent Butane") Then ValidRecipe = False
                            ' now is the mix % greater than the capabilities of the flow controller
                            If ValidRecipe Then
                                temp1 = ((100 - CSng(txtButnPercent.text)) / CSng(txtButnPercent.text)) * GramsPerHourToSlpm(CSng(txtLoadRate.text), sGramsPerLiter)
                                temp2 = 0.95 * Stn_AIO(DispStn, asNitrogenFlow).EuMax
                                If temp1 > temp2 Then
                                    ValidRecipe = False
                                    txtLoadRate.BackColor = EntryInvalid_BackColor
                                    txtButnPercent.BackColor = EntryInvalid_BackColor
                                    Message = "Required N2 flow(" & Format(temp1, "####0.000") & ") exceeds MFC range(" & Format(temp2, "###0.000") & ")."
                                    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
                                End If
                            End If
                        End If
                    End If
                Case STN_COMBO3_TYPE
                    ' future
                Case Else
                    ' nothing to do
            End Select
            
        End If
    End If
    
    If ValidRecipe Then
       
        ' Does the recipe actually DO anything
        If optNoPurge.Value = cYES _
            And optNoLoad.Value = cYES _
            And chkLeakCheck.Value = cNO Then
                ' Nothing to Do
                ValidRecipe = False
                optNoPurge.BackColor = EntryInvalid_BackColor
                optNoLoad.BackColor = EntryInvalid_BackColor
                chkLeakCheck.BackColor = EntryInvalid_BackColor
                Message = "Nothing to Do; No LeakCheck, No Purge, No Load"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
        End If
    End If
    

    ' Check Scales for ReadPort Errors
    If ValidRecipe Then
        ' Load stuff is O.K. &  Selected Scale Number are in Range
        If chkUseAuxScale = cYES Then
            ' Aux Scale
            If Not Scale_OK(CInt(txtAuxScaleNo)) Then
                ValidRecipe = False
                chkUseAuxScale.BackColor = EntryInvalid_BackColor
                txtAuxScaleNo.BackColor = EntryInvalid_BackColor
                Message = "Aux Scale " & txtAuxScaleNo.text & " has errors"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
        If chkPrimaryScale = cYES Then
            ' Primary Scale
            If Not Scale_OK(CInt(txtPrimaryScaleNo.text)) Then
                ValidRecipe = False
                chkPrimaryScale.BackColor = EntryInvalid_BackColor
                txtPrimaryScaleNo.BackColor = EntryInvalid_BackColor
                Message = "Primary Scale " & txtPrimaryScaleNo.text & " has errors"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
    End If

    ' Check for Exceeding Load Mass Limit
    If ValidRecipe Then
        ' Everything else is O.K.
        If USINGBUTANEMASSLIMIT Then
            ' using load mass limit
            ' calculate max mass to pass thru mfc
            If optLoadTime.Value = cYES Then
                ' LOADBYTIME
                maxrate = ValueFromText(txtLoadRate.text) + SysConfig.Tol_Btn_Flow
                maxmass = maxrate * (ValueFromText(txtLoadTime.text) / 60)
            ElseIf optWcm.Value = cYES Then
                ' LOADBYWC
                maxmass = StationCanister(DispStn, DispShift).WorkingCapacity * ValueFromText(txtWorkCapMult.text)
            ElseIf optLoadweight.Value = cYES Then
                ' LOADBYWEIGHT
                maxmass = ValueFromText(txtTargetWt.text)
            ElseIf optLoadBreakthrough.Value = cYES Then
                ' LOADBYBREAKTHRU
                maxmass = StationCanister(DispStn, DispShift).WorkingCapacity + ValueFromText(txtLoadBreakthrough.text)
            ElseIf optFIDBreakthrough.Value = cYES Then
                ' LOADBYFID
                maxmass = StationCanister(DispStn, DispShift).WorkingCapacity + (ValueFromText(txtFIDmg.text) / 1000)
            Else
                maxmass = 0
            End If
            ' check if limit is exceeded
            If (maxmass > (SysConfig.ButaneMassLimit * StationCanister(DispStn, DispShift).WorkingCapacity)) Then
                ValidRecipe = False
                If optLoadTime.Value = cYES Then
                    ' LOADBYTIME
                    txtLoadTime.BackColor = EntryInvalid_BackColor
                ElseIf optWcm.Value = cYES Then
                    ' LOADBYWC
                    txtWorkCapMult.BackColor = EntryInvalid_BackColor
                ElseIf optLoadweight.Value = cYES Then
                    ' LOADBYWEIGHT
                    txtTargetWt.BackColor = EntryInvalid_BackColor
                ElseIf optLoadBreakthrough.Value = cYES Then
                    ' LOADBYBREAKTHRU
                    txtLoadBreakthrough.BackColor = EntryInvalid_BackColor
                ElseIf optFIDBreakthrough.Value = cYES Then
                    ' LOADBYFID
                    txtFIDmg.BackColor = EntryInvalid_BackColor
                Else
                    txtLoadRate.BackColor = EntryInvalid_BackColor
                End If
                Message = "Recipe Load as selected will exceed the Load Mass Limit"
                lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
            End If
        End If
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

SetErrModule 90, 3
If UseLocalErrorHandler Then On Error GoTo localhandler
    Range_Check = True
    If (tcontrol.text = Empty) Then
        
        ' Empty Value
        Range_Check = False
        tcontrol.BackColor = EntryInvalid_BackColor
        Message = slabel & ":  Value is Empty!"
        lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
    
    ElseIf Not IsNumeric(tcontrol.text) Then
        
        ' Non-Numeric Value
        Range_Check = False
        tcontrol.BackColor = EntryInvalid_BackColor
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

Private Sub PurgeTime()
Dim tmpTime As Single
    If (optPurgeVolume.Value = cYES) Then
        ' PURGE BY VOLUME; calc duration
        If RecipeMode = STATIONMODE Then
        '    lblCalcPurgeUnits.Caption = "minutes"
            If IsNumeric(txtPurgeVolume) And IsNumeric(txtPurgeFlow) Then
                If StationCanister(DispStn, DispShift).WorkingVolume > 0 Then
                    If txtPurgeFlow > 0 And txtPurgeVolume > 0 Then
                        If Not USINGLINEVOLUME Then
                            tmpTime = ((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) / txtPurgeFlow)
        '                    txtCalcPurge = Format(((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) / txtPurgeFlow), "###0.0#")
        '                    txtCalcPurge.BackColor = frmHighlight.BackColor
        '                    txtCalcPurge.ForeColor = frmHighlight.ForeColor
                        Else
                            ' Using Line Volume
                            If IsNumeric(txtVentV) And IsNumeric(txtPurgeV) Then
                                tmpTime = (((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) _
                                    + txtVentV + txtPurgeV) / txtPurgeFlow)
        '                        txtCalcPurge = Format((((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) _
        '                            + txtVentV + txtPurgeV) / txtPurgeFlow), "###0.0#")
        '                        txtCalcPurge.BackColor = frmHighlight.BackColor
        '                        txtCalcPurge.ForeColor = frmHighlight.ForeColor
                            Else
                                Delay_Box "Invalid Line Volume", MSGDELAY, msgSHOW
                                tmpTime = 0
        '                        txtCalcPurge = "0"
        '                        txtCalcPurge.BackColor = EntryInvalid_BackColor
        '                        txtCalcPurge.ForeColor = BLACK
                            End If
                        End If
                    Else
                        If txtPurgeFlow <= 0 And txtPurgeVolume <= 0 Then
                            tmpTime = 0
        '                    txtCalcPurge = "0"
        '                    txtCalcPurge.BackColor = VERYPALEGRAY
        '                    txtCalcPurge.ForeColor = BLACK
                            txtPurgeFlow.BackColor = White
                            txtPurgeFlow.ForeColor = Black
                            txtPurgeVolume.BackColor = White
                            txtPurgeVolume.ForeColor = Black
                        ElseIf txtPurgeFlow <= 0 Then
            '                Delay_Box "No Purge Flow", MSGDELAY, msgSHOW
                            tmpTime = 0
        '                    txtCalcPurge = "0"
        '                    txtCalcPurge.BackColor = VERYPALEGRAY
        '                    txtCalcPurge.ForeColor = BLACK
                            txtPurgeFlow.BackColor = EntryInvalid_BackColor
                            txtPurgeFlow.ForeColor = Black
                        ElseIf txtPurgeVolume <= 0 Then
            '                Delay_Box "No Canister Volume", MSGDELAY, msgSHOW
                            tmpTime = 0
        '                    txtCalcPurge = "0"
        '                    txtCalcPurge.BackColor = VERYPALEGRAY
        '                    txtCalcPurge.ForeColor = BLACK
                            txtPurgeVolume.BackColor = EntryInvalid_BackColor
                            txtPurgeVolume.ForeColor = Black
                        End If
                    End If
                Else
                    Delay_Box "Save Canister Values first.", MSGDELAY, msgSHOW
                    tmpTime = 0
        '            txtCalcPurge = "0"
        '            txtCalcPurge.BackColor = VERYPALEGRAY
        '            txtCalcPurge.ForeColor = BLACK
                End If
            Else
                If Not IsNumeric(txtPurgeVolume) Then
            '        Delay_Box "Invalid Canister Volumes set to Zero.", MSGDELAY, msgSHOW
                    txtPurgeVolume = "0"
                    txtPurgeVolume.BackColor = EntryInvalid_BackColor
                    txtPurgeVolume.ForeColor = Black
                End If
                If Not IsNumeric(txtPurgeFlow) Then
            '        Delay_Box "Invalid Purge Flow Rate set to Zero.", MSGDELAY, msgSHOW
                    txtPurgeFlow = "0"
                    txtPurgeFlow.BackColor = EntryInvalid_BackColor
                    txtPurgeFlow.ForeColor = Black
                End If
                tmpTime = 0
        '        txtCalcPurge = "0"
        '        txtCalcPurge.BackColor = VERYPALEGRAY
        '        txtCalcPurge.ForeColor = BLACK
            End If
            txtPurgeTime.text = Format(tmpTime, "###0.0")
        End If
    ElseIf (optPurgeLiters.Value = cYES) Then
        ' PURGE BY LITERS; calc duration
        If IsNumeric(txtPurgeLiters) And IsNumeric(txtPurgeFlow) Then
            If (ValueFromText(txtPurgeFlow.text) > 0) And (ValueFromText(txtPurgeLiters.text) > 0) Then
                If Not USINGLINEVOLUME Then
                    tmpTime = ((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) / txtPurgeFlow)
'                    txtCalcPurge = Format(((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) / txtPurgeFlow), "###0.0#")
'                    txtCalcPurge.BackColor = frmHighlight.BackColor
'                    txtCalcPurge.ForeColor = frmHighlight.ForeColor
                Else
                    ' Using Line Volume
                    If IsNumeric(txtVentV) And IsNumeric(txtPurgeV) Then
                        tmpTime = ((txtPurgeLiters _
                            + txtVentV + txtPurgeV) / txtPurgeFlow)
'                        txtCalcPurge = Format((((StationCanister(DispStn, DispShift).WorkingVolume * txtPurgeVolume) _
'                            + txtVentV + txtPurgeV) / txtPurgeFlow), "###0.0#")
'                        txtCalcPurge.BackColor = frmHighlight.BackColor
'                        txtCalcPurge.ForeColor = frmHighlight.ForeColor
                    Else
                        Delay_Box "Invalid Line Volume", MSGDELAY, msgSHOW
                        tmpTime = 0
'                        txtCalcPurge = "0"
'                        txtCalcPurge.BackColor = EntryInvalid_BackColor
'                        txtCalcPurge.ForeColor = BLACK
                    End If
                End If
            Else
                If txtPurgeFlow <= 0 And txtPurgeLiters <= 0 Then
                    tmpTime = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = VERYPALEGRAY
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeFlow.BackColor = White
                    txtPurgeFlow.ForeColor = Black
                    txtPurgeLiters.BackColor = White
                    txtPurgeLiters.ForeColor = Black
                ElseIf txtPurgeFlow <= 0 Then
    '                Delay_Box "No Purge Flow", MSGDELAY, msgSHOW
                    tmpTime = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = VERYPALEGRAY
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeFlow.BackColor = EntryInvalid_BackColor
                    txtPurgeFlow.ForeColor = Black
                ElseIf txtPurgeLiters <= 0 Then
    '                Delay_Box "No Purge Liters", MSGDELAY, msgSHOW
                    tmpTime = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = VERYPALEGRAY
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeLiters.BackColor = EntryInvalid_BackColor
                    txtPurgeLiters.ForeColor = Black
                End If
            End If
        Else
            If Not IsNumeric(txtPurgeLiters) Then
        '        Delay_Box "Invalid Purge Liters set to Zero.", MSGDELAY, msgSHOW
                txtPurgeLiters = "0"
                txtPurgeLiters.BackColor = EntryInvalid_BackColor
                txtPurgeLiters.ForeColor = Black
            End If
            If Not IsNumeric(txtPurgeFlow) Then
        '        Delay_Box "Invalid Purge Flow Rate set to Zero.", MSGDELAY, msgSHOW
                txtPurgeFlow = "0"
                txtPurgeFlow.BackColor = EntryInvalid_BackColor
                txtPurgeFlow.ForeColor = Black
            End If
            tmpTime = 0
    '        txtCalcPurge = "0"
    '        txtCalcPurge.BackColor = VERYPALEGRAY
    '        txtCalcPurge.ForeColor = BLACK
        End If
        txtPurgeTime.text = Format(tmpTime, "###0.0")
    End If
End Sub

Private Sub PurgeVolume()
Dim tmpVolume As Single
If RecipeMode = STATIONMODE Then
'    lblCalcPurgeUnits.Caption = "volumes"
    If IsNumeric(txtPurgeTime) And IsNumeric(txtPurgeFlow) Then
        If StationCanister(DispStn, DispShift).WorkingVolume > 0 Then
            If txtPurgeFlow > 0 And txtPurgeTime > 0 Then
                If Not USINGLINEVOLUME Then
                    tmpVolume = ((txtPurgeTime * txtPurgeFlow) / (StationCanister(DispStn, DispShift).WorkingVolume))
'                    txtCalcPurge = Format(((txtPurgeTime * txtPurgeFlow) / (StationCanister(DispStn, DispShift).WorkingVolume)), "###0.0#")
'                    txtCalcPurge.BackColor = frmHighlight.BackColor
'                    txtCalcPurge.ForeColor = frmHighlight.ForeColor
                Else
                    ' Using Line Volume
                    If IsNumeric(txtVentV) And IsNumeric(txtPurgeV) Then
                        tmpVolume = (((txtPurgeFlow * txtPurgeTime) _
                            - txtVentV - txtPurgeV) / StationCanister(DispStn, DispShift).WorkingVolume)
'                        txtCalcPurge = Format((((txtPurgeFlow * txtPurgeTime) _
'                            - txtVentV - txtPurgeV) / StationCanister(DispStn, DispShift).WorkingVolume), "###0.0#")
'                        txtCalcPurge.BackColor = frmHighlight.BackColor
'                        txtCalcPurge.ForeColor = frmHighlight.ForeColor
                    Else
                        Delay_Box "Invalid Line Volume", MSGDELAY, msgSHOW
                        tmpVolume = 0
'                        txtCalcPurge = "0"
'                        txtCalcPurge.BackColor = EntryInvalid_BackColor
'                        txtCalcPurge.ForeColor = BLACK
                    End If
                End If
            Else
                If txtPurgeFlow <= 0 And txtPurgeTime <= 0 Then
                    tmpVolume = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = Common_BackColor
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeFlow.BackColor = White
                    txtPurgeFlow.ForeColor = Black
                    txtPurgeTime.BackColor = White
                    txtPurgeTime.ForeColor = Black
                ElseIf txtPurgeFlow <= 0 Then
    '                Delay_Box "No Purge Flow", MSGDELAY, msgSHOW
                    tmpVolume = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = Common_BackColor
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeFlow.BackColor = EntryInvalid_BackColor
                    txtPurgeFlow.ForeColor = Black
                ElseIf txtPurgeTime <= 0 Then
    '                Delay_Box "No Time", MSGDELAY, msgSHOW
                    tmpVolume = 0
'                    txtCalcPurge = "0"
'                    txtCalcPurge.BackColor = Common_BackColor
'                    txtCalcPurge.ForeColor = BLACK
                    txtPurgeTime.BackColor = EntryInvalid_BackColor
                    txtPurgeTime.ForeColor = Black
                End If
            End If
        Else
            Delay_Box "Save Canister Values first.", MSGDELAY, msgSHOW
            tmpVolume = 0
'            txtCalcPurge = "0"
'            txtCalcPurge.BackColor = Common_BackColor
'            txtCalcPurge.ForeColor = BLACK
        End If
    Else
        If Not IsNumeric(txtPurgeTime) Then
    '        Delay_Box "Invalid Purge Time; set to Zero.", MSGDELAY, msgSHOW
            txtPurgeTime = "0"
            txtPurgeTime.BackColor = EntryInvalid_BackColor
            txtPurgeTime.ForeColor = Black
        End If
        If Not IsNumeric(txtPurgeFlow) Then
    '        Delay_Box "Invalid Purge Flow Rate; set to Zero.", MSGDELAY, msgSHOW
            txtPurgeFlow = "0"
            txtPurgeFlow.BackColor = EntryInvalid_BackColor
            txtPurgeFlow.ForeColor = Black
        End If
        tmpVolume = 0
'        txtCalcPurge = "0"
'        txtCalcPurge.BackColor = Common_BackColor
'        txtCalcPurge.ForeColor = BLACK
    End If
    txtPurgeVolume.text = Format(tmpVolume, "###0.0#")
End If
End Sub

Private Sub chkADF_Heater_Click()
    chkADF_Heater.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub chkLeakAux_Click()
    chkLeakCheck.BackColor = frmNotHighlight.BackColor
    chkLeakPrimary.BackColor = frmNotHighlight.BackColor
    chkLeakAux.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub chkLeakCheck_Click()
    chkLeakCheck.BackColor = frmNotHighlight.BackColor
    chkLeakPrimary.BackColor = frmNotHighlight.BackColor
    chkLeakAux.BackColor = frmNotHighlight.BackColor
    chkPauseAfterLeak.BackColor = frmNotHighlight.BackColor
    txtPauseLeakTime.BackColor = frmNotHighlight.BackColor
    If (chkLeakCheck.Value = cNO) Then
        chkLeakPrimary.Value = cNO
        chkLeakAux.Value = cNO
        chkPauseAfterLeak.Value = cNO
        txtPauseLeakTime.text = "0"
    End If
End Sub

Private Sub chkLiveFuel_Click()
    chkLiveFuel.BackColor = frmNotHighlight.BackColor
    Select Case RecipeMode
        Case MASTERMODE
            Select Case chkLiveFuel.Value
                Case cYES
                    chkOrvrMfc.Value = cNO
                    chkOrvrMfc.Visible = IIf(systemhasORVR2, True, False)
                    txtButnPercent.text = "0"
                    txtButnPercent.Visible = False
                    lblButnPercent.Visible = False
                    lblButnPercentUnits.Visible = False
                    chkLoadRatePID.Visible = True
                Case cNO
                    txtNitrogenFlow.text = "0"
                    txtLiveFuelChgFreq.text = "0"
                    txtADF_HeaterSP.text = "0"
                    chkLiveFuelChgAuto.Value = IIf(systemhasLIVEFUEL, cYES, cNO)
                    chkADF_Heater.Value = IIf(systemhasADF_HEATER, cYES, cNO)
                    chkOrvrMfc.Visible = IIf(systemhasORVR2, True, False)
                    txtButnPercent.Visible = IIf(systemhasBUTANE, True, False)
                    lblButnPercent.Visible = IIf(systemhasBUTANE, True, False)
                    lblButnPercentUnits.Visible = IIf(systemhasBUTANE, True, False)
                    chkLoadRatePID.Value = cNO
                    chkLoadRatePID.Visible = False
            End Select
        Case STATIONMODE
            Select Case STN_INFO(DispStn).Type
                Case STN_REGULAR_TYPE, STN_ORVR_TYPE, STN_ORVR2_TYPE
                    chkLiveFuel.Value = cNO
                    chkLoadRatePID.Value = cNO
                Case STN_LIVEFUEL_TYPE
'                    chkLiveFuel.Value = cYES
                Case STN_LIVEREG_TYPE
                    Select Case chkLiveFuel.Value
                        Case cYES
                            txtButnPercent.Visible = False
                            lblButnPercent.Visible = False
                            lblButnPercentUnits.Visible = False
                            chkLoadRatePID.Visible = True
                            chkLiveFuelChgAuto.Value = IIf((STN_INFO(DispStn).ADF_TANKTYPE = 90), cNO, cYES)
                        Case cNO
                            txtButnPercent.Visible = True
                            lblButnPercent.Visible = True
                            lblButnPercentUnits.Visible = True
                            chkLoadRatePID.Value = cNO
                            chkLoadRatePID.Visible = False
                            chkLiveFuelChgAuto.Value = cNO
                            txtNitrogenFlow.text = "0"
                            txtLiveFuelChgFreq.text = "0"
                            txtADF_HeaterSP.text = "0"
                            chkLiveFuelChgAuto.Value = cNO
                            chkADF_Heater.Value = cNO
                    End Select
                Case STN_LIVEORVR2_TYPE
                    Select Case chkLiveFuel.Value
                        Case cYES
                            txtButnPercent.Visible = False
                            lblButnPercent.Visible = False
                            lblButnPercentUnits.Visible = False
                            chkLoadRatePID.Visible = True
                            chkLiveFuelChgAuto.Value = IIf((STN_INFO(DispStn).ADF_TANKTYPE = 90), cNO, cYES)
                        Case cNO
                            Select Case chkOrvrMfc.Value
                                Case cYES
                                    txtButnPercent.Visible = True
                                    lblButnPercent.Visible = True
                                    lblButnPercentUnits.Visible = True
                                    chkLoadRatePID.Value = cNO
                                    chkLoadRatePID.Visible = False
                                    txtNitrogenFlow.text = "0"
                                    txtLiveFuelChgFreq.text = "0"
                                    txtADF_HeaterSP.text = "0"
'                                    chkLiveFuelChgAuto.Value = cNO
                                    chkADF_Heater.Value = cNO
                                Case cNO
                                    txtButnPercent.Visible = True
                                    lblButnPercent.Visible = True
                                    lblButnPercentUnits.Visible = True
                                    chkLoadRatePID.Value = cNO
                                    chkLoadRatePID.Visible = False
                                    txtNitrogenFlow.text = "0"
                                    txtLiveFuelChgFreq.text = "0"
                                    txtADF_HeaterSP.text = "0"
'                                    chkLiveFuelChgAuto.Value = cNO
                                    chkADF_Heater.Value = cNO
                            End Select
                    End Select
                Case STN_COMBO3_TYPE
                    ' future
                Case Else
                    ' nothing to do
            End Select
    End Select
End Sub

Private Sub chkLiveFuelChgAuto_Click()
    chkLiveFuelChgAuto.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub chkLeakPrimary_Click()
    chkLeakCheck.BackColor = frmNotHighlight.BackColor
    chkLeakPrimary.BackColor = frmNotHighlight.BackColor
    chkLeakAux.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub optPauseAfterLoad_Click()
    optPauseAfterLoad.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub optPauseAfterLoadForOper_Click()
    optPauseAfterLoadForOper.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub optPauseAfterPurge_Click()
    optPauseAfterPurge.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub optPauseAfterPurgeForOper_Click()
    optPauseAfterPurgeForOper.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub chkPrimaryScale_Click()
    chkPrimaryScale.BackColor = frmNotHighlight.BackColor
    txtPrimaryScaleNo.BackColor = txtNotHighlight.BackColor
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        chkPrimaryScale.Value = cYES
    End If
End Sub

Private Sub chkPurgeAuxCan_Click()
    chkPurgeAuxCan.BackColor = frmNotHighlight.BackColor
End Sub

Private Sub chkUseAuxScale_Click()
    chkUseAuxScale.BackColor = frmNotHighlight.BackColor
    txtAuxScaleNo.BackColor = txtNotHighlight.BackColor
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        chkUseAuxScale.Value = cYES
    End If
End Sub

Private Sub cmdPrint_Click()
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    ' Print Current Recipe
    lblMessage.Caption = ""
    Set pbCapture.Picture = CaptureForm(Me)
    PrintPictureToFitPage Printer, pbCapture.Picture
    Printer.EndDoc
    Set pbCapture.Picture = Nothing
    lblMessage.Font.Size = 9.5
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = "Recipe sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdPrintAll_Click()
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    Select Case RecipeMode
        Case MASTERMODE
            ' Print All Master Recipes
            lblMessage.Caption = ""
            Print_AllMasters
            lblMessage.Font.Size = 9.5
            lblMessage.ForeColor = Message_ForeColor
            lblMessage.Caption = "Master Recipes sent to" & vbCrLf & PRINTERNAME
        Case STATIONMODE
            ' Print All Station Recipes
            lblMessage.Caption = ""
            Print_AllStationRecipes
            lblMessage.Font.Size = 9.5
            lblMessage.ForeColor = Message_ForeColor
            lblMessage.Caption = "Station Recipes sent to" & vbCrLf & PRINTERNAME
    End Select
End Sub

Private Sub cmdPurgeProfile_Click()
    Select Case RecipeMode
        Case MASTERMODE
'            frmSearchProf.Show
'            frmSearchProf.ChgSelectionDestination (CInt(profdestRecipe))
            frmPurgeProfile.Show
            frmPurgeProfile.ChgProfileMode (CInt(MASTERMODE))
            frmPurgeProfile.InitProfile
        Case STATIONMODE
            frmPurgeProfile.Show
            frmPurgeProfile.ChgProfileMode (CInt(STATIONMODE))
            frmPurgeProfile.InitProfile
    End Select
End Sub

Private Sub cmdPurgeWizard_Click()
    frmWizardPrg.Show
End Sub

Private Sub cmdRestore_Click()
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    RecipeDisplay_ByStnShift
End Sub

Private Sub cmdReturn_Click()
    ExitScreen
End Sub

Private Sub cmdSave_Click()
    SaveRecipe
End Sub

Public Sub SaveRecipe()
SetErrModule 90, 2
If UseLocalErrorHandler Then On Error GoTo localhandler
Dim sMsg As String
Dim rcpDur As Single

    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    
    Select Case RecipeMode
        Case MASTERMODE
            ' master
            If CheckPass("O", False) Then
                lblMessage.Caption = vbCrLf
                Reset_BackColors
                If ValidRecipe Then
                    Reset_BackColors
                    ScreenToDspRcp              ' Copy screen data to Recipe Array
                    ' Save Master Recipe Information
                    SaveMasterRcp CInt(DspRecipe.Number)
                    ' Save Remote Master Canister Information
                    If USINGREMCANLOAD Then
                        ' open master canister / recipe database
                        Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
                        ' open remote database
                        OpenConnToRemoteDb
                        ' update Remote Master Recipe Information
                        UpdateRemoteRecipes
                        ' close remote database
                        CloseConnToRemoteDb
                    End If
                    Chgs = False
                    lblMessage.ForeColor = Message_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "New Recipe Settings Saved" & vbCrLf
                Else
                    lblMessage.ForeColor = Alarm_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "Recipe Settings Not Saved" & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & "Try again after correcting the errors." & vbCrLf
                    Beep
                    Beep
                    Beep
                End If
            Else
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Insufficient Access" & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "The current user is not authorized to save Master Recipes" & vbCrLf
                Beep
                Beep
                Beep
            End If
        
        Case STATIONMODE
            ' station
            If StationControl(DispStn, DispShift).Mode = VBIDLE Then
                lblMessage.Caption = vbCrLf
                Reset_BackColors
                If ValidRecipe Then
                    Reset_BackColors
                    ' Copy screen data to Station Recipe
                    ScreenToDspRcp
                    ' Update Station Recipe Name & (Master Recipe)Number
                    StationRecipe(DispStn, DispShift) = DspRecipe
                    StationRecipe(DispStn, DispShift).Number = IIf(Chgs, CInt(0), DispRcp)
                    ' Update other Station Recipe descriptors
                    UpdateStnRcpDsc DispStn, DispShift
                    ' save station recipes
                    Save_StationRecipes
                    ' estimated Recipe Duration
                    rcpDur = EstimatedRcpDuration(StationRecipe(DispStn, DispShift), StationCanister(DispStn, DispShift), StationProfile(DispStn, DispShift))
                    sMsg = "Estimated Recipe Duration is " & DurationDescription(rcpDur)
                    ' update job sequence duration, if required
                    ' recipe 0 means use existing station recipe
                    If StationSequence(DispStn, DispShift).CourseData(1).RecipeNumber = 0 Then
                        ' using existing Station Recipe
                        ' if only 1 course then use recipe duration for job sequence duration
                        Select Case StationSequence(DispStn, DispShift).NumCourses
                            Case 1
                                ' estimated Job Sequence Duration
                                StationSequence(DispStn, DispShift).EstSeqDuration = rcpDur
                                StationSequence(DispStn, DispShift).EstSeqDurDesc = sMsg
                            Case Else
                                ' nothing to do
                        End Select
                    End If
                    ' update live fuel parameters for adf
                    LiveFuel_Update DispStn, DispShift
                    ' recipe saved
                    lblMessage.ForeColor = Message_ForeColor
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    Select Case NR_SHIFT
                        Case 1
                            lblMessage.Caption = lblMessage.Caption & "New Recipe Settings Saved to Station #" + Format(DispStn, "0")
                        Case 2
                            lblMessage.Caption = lblMessage.Caption & "New Recipe Settings Saved to Station #" + Format(DispStn, "0") + " / Shift #" + Format(DispShift, "0")
                    End Select
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                    lblMessage.Caption = lblMessage.Caption & sMsg
                    lblMessage.Caption = lblMessage.Caption & vbCrLf
                End If
            Else
                lblMessage.ForeColor = Alarm_ForeColor
                lblMessage.Caption = lblMessage.Caption & vbCrLf
                lblMessage.Caption = lblMessage.Caption & "Recipe Settings Not Saved" & vbCrLf
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

Private Sub cmdDown_Click()
  DispRcp = IIf(DispRcp < 2, NR_RCP, DispRcp - 1)
  RecipeDisplay_ByNum
End Sub

Private Sub cmdUp_Click()
  DispRcp = IIf(DispRcp > NR_RCP - 1, 1, DispRcp + 1)
  RecipeDisplay_ByNum
End Sub

Private Sub cmdPgDn_Click()
  DispRcp = IIf(DispRcp < 11, NR_RCP, DispRcp - 10)
  RecipeDisplay_ByNum
End Sub

Private Sub cmdPgUp_Click()
  DispRcp = IIf(DispRcp > NR_RCP - 10, 1, DispRcp + 10)
  RecipeDisplay_ByNum
End Sub

Private Sub Form_Activate()
    Form_Center Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        ExitScreen
    End If
End Sub

Private Sub Form_Load()

    KeyPreview = True
    UpdateRecipeScreen
    ' open canister / recipe database
    Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
    
End Sub

Private Sub UpdateRecipeScreen()
Dim tmpColor As Long
Dim flag As Boolean
Dim Idx As Integer
Dim sTxt As String
    
    txtStartAtDate.ToolTipText = "Enter Start Time as YYYY-MM-DD hh:mm"

    ' Set Title Foreground colors
    tmpColor = TitlesLabel_ForeColor
        frmCycle.ForeColor = tmpColor
    tmpColor = TitlesData_Forecolor
        frmEnd.ForeColor = tmpColor
        frmStart.ForeColor = tmpColor
        frmPurge.ForeColor = tmpColor
        frmLoad.ForeColor = tmpColor
        frmLeakCheck.ForeColor = tmpColor
        frmAuxOutputs.ForeColor = tmpColor
    tmpColor = Titles_ForeColor
        frmResources.ForeColor = tmpColor
        frmLineVolume.ForeColor = tmpColor
        frmLiveFuel.ForeColor = tmpColor
    tmpColor = TitlesData_Forecolor
        pnlRecipe.ForeColor = tmpColor
        txtRecipeName.ForeColor = tmpColor
        pnlDispRcpNum.ForeColor = tmpColor
    
    ' Reset all the backgrounds
    Reset_BackColors

    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = ""

    optPauseAfterPurge.Enabled = True
    txtPausePurgeTime.Enabled = True
    optPauseAfterLoad.Enabled = True
    txtPauseLoadTime.Enabled = True
    chkPauseAfterLeak.Enabled = True
    txtPauseLeakTime.Enabled = True
    optPauseAfterLoadForOper.Enabled = True
    optPauseAfterPurgeForOper.Enabled = True
    optNoPauseAfterLoad.Enabled = True
    optNoPauseAfterPurge.Enabled = True
'    If USINGLOADTIMELIMIT Then
'        lblMaxLoadTime.Visible = True
'        txtMaxLoadTime.Visible = True
'    Else
        lblMaxLoadTime.Visible = False
        txtMaxLoadTime.Visible = False
'    End If
    
    If USINGAUXLEAKCHECK Then
        chkLeakPrimary.Visible = True
        chkLeakAux.Visible = True
    Else
        chkLeakPrimary.Visible = False
        chkLeakAux.Visible = False
        chkLeakPrimary.Value = cYES
        chkLeakAux.Value = cNO
    End If
    
    ' Purge Cans in Series only allowed with Vacuum Purge
    chkPurgeCansInSeries.Visible = IIf(((Not SysConfig.PosPressPurge) And USINGPURGESERIES), True, False)
    
    ' Purge Oven option
    Select Case RecipeMode
        Case MASTERMODE
            flag = IIf((USINGPURGEOVEN), True, False)
        Case STATIONMODE
            flag = IIf((USINGPURGEOVEN And STN_INFO(DispStn).USINGPURGEOVEN), True, False)
    End Select
    chkUsePurgeOven.Visible = flag
    lblPurgeOvenUnits.Visible = flag
    txtPurgeOvenSP.Visible = flag
    
    If NR_SCALES > 1 Then
        txtPrimaryScaleNo.ToolTipText = "Enter Scale# from 1 to " & Format(NR_SCALES, "#0")
        txtAuxScaleNo.ToolTipText = "Enter Scale# from 1 to " & Format(NR_SCALES, "#0")
    Else
        txtPrimaryScaleNo.ToolTipText = "Enter Scale# 1"
        txtAuxScaleNo.ToolTipText = "Enter Scale# 1"
    End If
    
    ' Line Volume
    If Not USINGLINEVOLUME Then
        frmLineVolume.Visible = False
    Else
        frmLineVolume.Visible = True
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
    End If
    
    'Live Fuel
    If systemhasLIVEFUEL Then
        chkLoadRatePID.Visible = True
        frmLiveFuel.Visible = True
        chkLiveFuel.Enabled = True
        chkLiveFuel.Visible = True
        lblNitrogenFlow.Visible = True
        txtNitrogenFlow.Enabled = True
        txtNitrogenFlow.Visible = True
        lblLiveFuelChgFreq.Visible = True
        lblLiveFuelChgFreq2.Visible = True
        txtLiveFuelChgFreq.Enabled = True
        txtLiveFuelChgFreq.Visible = True
        If systemhasAUTODRAINFILL Then
           If (STN_INFO(DispStn).ADF_TANKTYPE > 10) And (STN_INFO(DispStn).ADF_TANKTYPE <= 20) Then
                ' intank heater
                chkLiveFuelChgAuto.Enabled = True
                chkLiveFuelChgAuto.Visible = True
                chkADF_Heater.Enabled = True
                chkADF_Heater.Visible = True
                lblADF_HeaterSP.Visible = True
                If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                txtADF_HeaterSP.Enabled = True
                txtADF_HeaterSP.Visible = True
                If USINGC Then txtADF_HeaterSP.ToolTipText = "15 to 50 deg C"
                If USINGF Then txtADF_HeaterSP.ToolTipText = "60 to 120 deg F"
            ElseIf (STN_INFO(DispStn).ADF_TANKTYPE = 90) Then
                ' waterbath heater
                chkLiveFuelChgAuto.Visible = False
                chkADF_Heater.Enabled = True
                chkADF_Heater.Visible = True
                lblADF_HeaterSP.Visible = True
                If USINGC Then lblADF_HeaterSP.Caption = "deg C"
                If USINGF Then lblADF_HeaterSP.Caption = "deg F"
                txtADF_HeaterSP.Enabled = True
                txtADF_HeaterSP.Visible = True
                Select Case RecipeMode
                    Case MASTERMODE
                        Idx = SysConfig.WaterBathControl
                    Case STATIONMODE
                        Idx = StationConfig(DispStn, DispShift).WaterBathControl
                End Select
                Select Case Idx
                    Case wbDirect
                        sTxt = "WaterBath SetPoint from "
                    Case wbFuelTemp
                        sTxt = "LiveFuel Fuel SetPoint from "
                    Case wbVaporTemp
                        sTxt = "LiveFuel Vapor SetPoint from "
                End Select
                If USINGC Then txtADF_HeaterSP.ToolTipText = sTxt & Format(WB_AIO.EuMin, "###0.0##") & " to " & Format(WB_AIO.EuMax, "###0.0##") & " deg C"
                If USINGF Then txtADF_HeaterSP.ToolTipText = sTxt & Format(DegCtoF(WB_AIO.EuMin), "###0.0##") & " to " & Format(DegCtoF(WB_AIO.EuMax), "###0.0##") & " deg F"
            Else
                chkLiveFuelChgAuto.Visible = True
                chkADF_Heater.Enabled = False
                chkADF_Heater.Visible = False
                lblADF_HeaterSP.Visible = False
                txtADF_HeaterSP.Enabled = False
                txtADF_HeaterSP.Visible = False
            End If
        Else
            chkLiveFuelChgAuto.Enabled = False
            chkLiveFuelChgAuto.Visible = False
            chkADF_Heater.Enabled = False
            chkADF_Heater.Visible = False
            lblADF_HeaterSP.Visible = False
            txtADF_HeaterSP.Enabled = False
            txtADF_HeaterSP.Visible = False
        End If
    Else
        chkLoadRatePID.Visible = False
        frmLiveFuel.Visible = False
    End If
    
    ' orvr mfc's
    chkOrvrMfc.Visible = IIf(systemhasORVR2, True, False)
    
    ' aux outputs
    frmAuxOutputs.Top = OutOfSight
    If USING_AUX_OUTPUTS Then
        chkAuxOutputs.Visible = True
        cmdCfgAuxOutputs.Visible = True
    Else
        chkAuxOutputs.Visible = False
        cmdCfgAuxOutputs.Visible = False
    End If
    chkAuxLoad(1).Caption = DESC_AUX_OUTPUT1
    chkAuxLoad(2).Caption = DESC_AUX_OUTPUT2
    chkAuxLoad(3).Caption = DESC_AUX_OUTPUT3
    chkAuxLoad(4).Caption = DESC_AUX_OUTPUT4
    chkAuxPurge(1).Caption = DESC_AUX_OUTPUT1
    chkAuxPurge(2).Caption = DESC_AUX_OUTPUT2
    chkAuxPurge(3).Caption = DESC_AUX_OUTPUT3
    chkAuxPurge(4).Caption = DESC_AUX_OUTPUT4

    ' Recipe Information Panel
    frmRcpInfo.Top = OutOfSight
    
    ' authorized to Save Master Recipes ?
    cmdSave.Visible = IIf(CheckPass("O", False), True, False)
    
    ' show Restore Station Recipe ?
    cmdRestore.Visible = IIf(RecipeMode = STATIONMODE, True, False)
        
End Sub

Private Sub optPurgeLiters_Click()
    If optPurgeLiters.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeAuxOnly.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        optPurgeLiters.BackColor = frmHighlight.BackColor
        optPurgeLiters.ForeColor = frmHighlight.ForeColor
        txtPurgeLiters.BackColor = frmHighlight.ForeColor
        txtPurgeLiters.ForeColor = frmHighlight.BackColor
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
        optTargetContinuous.Enabled = True
        optTargetPurgePauseRepeat.Enabled = True
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = True
        lblTargetPauseUnits.Enabled = True
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = True
        txtTargetPause.Enabled = True
    Else
        optPurgeLiters.BackColor = frmNotHighlight.BackColor
        optPurgeLiters.ForeColor = frmNotHighlight.ForeColor
        txtPurgeLiters.BackColor = txtNotHighlight.BackColor
        txtPurgeLiters.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub tabsCycleType_Click()
    Select Case tabsCycletype.SelectedItem
        Case "Purge - Load"
            optEndCycles.Caption = "End after             Purge / Load Cycles"
        Case "Load - Purge"
            optEndCycles.Caption = "End after             Load / Purge Cycles"
    End Select
End Sub

Private Sub optEndCycles_Click()
    If optEndCycles.Value = cYES Then
        optEndWeightChange.Value = cNO
        optUpdateCanWc.Value = cNO
        optEndCycles.BackColor = frmHighlight.BackColor
        optEndCycles.ForeColor = frmHighlight.ForeColor
        txtPFCycle.BackColor = frmHighlight.ForeColor
        txtPFCycle.ForeColor = frmHighlight.BackColor
    Else
        optEndCycles.BackColor = frmNotHighlight.BackColor
        optEndCycles.ForeColor = frmNotHighlight.ForeColor
        txtPFCycle.BackColor = txtNotHighlight.BackColor
        txtPFCycle.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optEndWeightChange_Click()
    If optEndWeightChange.Value = cYES Then
        optEndCycles.Value = cNO
        optEndWeightChange.BackColor = frmHighlight.BackColor
        optEndWeightChange.ForeColor = frmHighlight.ForeColor
        txtWeightChangeTol.BackColor = frmHighlight.BackColor
        txtWeightChangeTol.ForeColor = frmHighlight.ForeColor
        txtConsecutiveCycles.BackColor = frmHighlight.BackColor
        txtConsecutiveCycles.ForeColor = frmHighlight.ForeColor
        txtMaximumCycles.BackColor = frmHighlight.BackColor
        txtMaximumCycles.ForeColor = frmHighlight.ForeColor
        txtMinimumCycles.BackColor = frmHighlight.BackColor
        txtMinimumCycles.ForeColor = frmHighlight.ForeColor
        optUpdateCanWc.BackColor = frmHighlight.BackColor
        optUpdateCanWc.ForeColor = frmHighlight.ForeColor
    Else
        optEndWeightChange.BackColor = frmNotHighlight.BackColor
        optEndWeightChange.ForeColor = frmNotHighlight.ForeColor
        txtWeightChangeTol.BackColor = txtNotHighlight.BackColor
        txtWeightChangeTol.ForeColor = txtNotHighlight.ForeColor
        txtConsecutiveCycles.BackColor = txtNotHighlight.BackColor
        txtConsecutiveCycles.ForeColor = txtNotHighlight.ForeColor
        txtMaximumCycles.BackColor = txtNotHighlight.BackColor
        txtMaximumCycles.ForeColor = txtNotHighlight.ForeColor
        txtMinimumCycles.BackColor = txtNotHighlight.BackColor
        txtMinimumCycles.ForeColor = txtNotHighlight.ForeColor
        optUpdateCanWc.BackColor = frmNotHighlight.BackColor
        optUpdateCanWc.ForeColor = frmNotHighlight.ForeColor
    End If
End Sub

Private Sub optFIDBreakthrough_Click()
    If optFIDBreakthrough.Value = cYES Then
        txtLoadTime.text = "0"
        txtWorkCapMult.text = "0"
        txtEPAFill.text = "0"
        txtTargetWt.text = "0"
        txtLoadBreakthrough.text = "0"
        optLoadTime.Value = cNO
        optWcm.Value = cNO
        optLoadweight.Value = cNO
        optLoadBreakthrough.Value = cNO
        optNoLoad.Value = cNO
        optFIDBreakthrough.BackColor = frmHighlight.BackColor
        optFIDBreakthrough.ForeColor = frmHighlight.ForeColor
        txtFIDmg.BackColor = frmHighlight.ForeColor
        txtFIDmg.ForeColor = frmHighlight.BackColor
    Else
        optFIDBreakthrough.BackColor = frmNotHighlight.BackColor
        optFIDBreakthrough.ForeColor = frmNotHighlight.ForeColor
        txtFIDmg.BackColor = txtNotHighlight.BackColor
        txtFIDmg.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optLoadBreakthrough_Click()
    If optLoadBreakthrough.Value = cYES Then
        txtLoadTime.text = "0"
        txtWorkCapMult.text = "0"
        txtEPAFill.text = "0"
        txtTargetWt.text = "0"
        txtFIDmg.text = "0"
        optLoadTime.Value = cNO
        optWcm.Value = cNO
        optLoadweight.Value = cNO
        optNoLoad.Value = cNO
        optFIDBreakthrough.Value = cNO
        optLoadBreakthrough.BackColor = frmHighlight.BackColor
        optLoadBreakthrough.ForeColor = frmHighlight.ForeColor
        txtLoadBreakthrough.BackColor = frmHighlight.ForeColor
        txtLoadBreakthrough.ForeColor = frmHighlight.BackColor
    Else
        optLoadBreakthrough.BackColor = frmNotHighlight.BackColor
        optLoadBreakthrough.ForeColor = frmNotHighlight.ForeColor
        txtLoadBreakthrough.BackColor = txtNotHighlight.BackColor
        txtLoadBreakthrough.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optLoadTime_Click()
    If optLoadTime.Value = cYES Then
        txtWorkCapMult.text = "0"
        txtEPAFill.text = "0"
        txtTargetWt.text = "0"
        txtLoadBreakthrough.text = "0"
        txtFIDmg.text = "0"
        optNoLoad.Value = cNO
        optWcm.Value = cNO
        optLoadweight.Value = cNO
        optLoadBreakthrough.Value = cNO
        optFIDBreakthrough.Value = cNO
        optLoadTime.BackColor = frmHighlight.BackColor
        optLoadTime.ForeColor = frmHighlight.ForeColor
        txtLoadTime.BackColor = frmHighlight.ForeColor
        txtLoadTime.ForeColor = frmHighlight.BackColor
    Else
        optLoadTime.BackColor = frmNotHighlight.BackColor
        optLoadTime.ForeColor = frmNotHighlight.ForeColor
        txtLoadTime.BackColor = txtNotHighlight.BackColor
        txtLoadTime.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optLoadweight_Click()
    If optLoadweight.Value = cYES Then
        txtLoadTime.text = "0"
        txtWorkCapMult.text = "0"
        txtEPAFill.text = "0"
        txtLoadBreakthrough.text = "0"
        txtFIDmg.text = "0"
        optLoadTime.Value = cNO
        optWcm.Value = cNO
        optNoLoad.Value = cNO
        optLoadBreakthrough.Value = cNO
        optFIDBreakthrough.Value = cNO
        optLoadweight.BackColor = frmHighlight.BackColor
        optLoadweight.ForeColor = frmHighlight.ForeColor
        txtTargetWt.BackColor = frmHighlight.ForeColor
        txtTargetWt.ForeColor = frmHighlight.BackColor
    Else
        optLoadweight.BackColor = frmNotHighlight.BackColor
        optLoadweight.ForeColor = frmNotHighlight.ForeColor
        txtTargetWt.BackColor = txtNotHighlight.BackColor
        txtTargetWt.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optNoLoad_Click()
    If optNoLoad.Value = cYES Then
        txtLoadTime.text = "0"
        txtWorkCapMult.text = "0"
        txtEPAFill.text = "0"
        txtTargetWt.text = "0"
        txtLoadBreakthrough.text = "0"
        txtFIDmg.text = "0"
        txtNitrogenFlow.text = "0"
        txtLiveFuelChgFreq.text = "0"
        txtADF_HeaterSP.text = "0"
        txtLoadRate.text = "0"
        txtButnPercent.text = "0"
        txtPauseLoadTime.text = "0"
        optPauseAfterLoad.Value = False
        optPauseAfterLoadForOper.Value = False
        chkLiveFuel.Value = cNO
        chkLiveFuelChgAuto.Value = cNO
        chkADF_Heater.Value = cNO
        optLoadTime.Value = cNO
        optWcm.Value = cNO
        optLoadweight.Value = cNO
        optLoadBreakthrough.Value = cNO
        optFIDBreakthrough.Value = cNO
        optNoLoad.BackColor = frmHighlight.BackColor
        optNoLoad.ForeColor = frmHighlight.ForeColor
    Else
        optNoLoad.BackColor = frmNotHighlight.BackColor
        optNoLoad.ForeColor = frmNotHighlight.ForeColor
    End If
End Sub

Private Sub optNoPurge_Click()
    If optNoPurge.Value = cYES Then
        txtPurgeTime.text = "0"
        txtPurgeAuxOnly.text = "0"
        txtPurgeProfile.text = "0"
        txtPurgeVolume.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        txtTargetPurge.text = "0"
        txtTargetPause.text = "0"
        txtPurgeFlow.text = "0"
        txtPausePurgeTime.text = "0"
        chkPurgeAuxCan.Value = cNO
        chkPurgeCansInSeries.Value = cNO
        txtPurgeTime.text = "0"
        optPauseAfterPurge.Value = False
        optPurgeTime.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        optNoPurge.BackColor = frmHighlight.BackColor
        optNoPurge.ForeColor = frmHighlight.ForeColor
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
        optTargetContinuous.Enabled = False
        optTargetPurgePauseRepeat.Enabled = False
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = False
        lblTargetPauseUnits.Enabled = False
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = False
        txtTargetPause.Enabled = False
    Else
        optNoPurge.BackColor = frmNotHighlight.BackColor
        optNoPurge.ForeColor = frmNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeTime_Click()
    If optPurgeTime.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        If IsNumeric(txtPurgeTime.text) Then txtPurgeTime.text = Format(CInt(txtPurgeTime.text), "####0")
        txtPurgeVolume.text = "0"
        txtPurgeAuxOnly.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        optPurgeTime.BackColor = frmHighlight.BackColor
        optPurgeTime.ForeColor = frmHighlight.ForeColor
        txtPurgeTime.BackColor = frmHighlight.ForeColor
        txtPurgeTime.ForeColor = frmHighlight.BackColor
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
        optTargetContinuous.Enabled = False
        optTargetPurgePauseRepeat.Enabled = False
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = False
        lblTargetPauseUnits.Enabled = False
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = False
        txtTargetPause.Enabled = False
    Else
        optPurgeTime.BackColor = frmNotHighlight.BackColor
        optPurgeTime.ForeColor = frmNotHighlight.ForeColor
        txtPurgeTime.BackColor = txtNotHighlight.BackColor
        txtPurgeTime.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeUndo_Click()
    If optPurgeUndo.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeUndo.BackColor = frmHighlight.BackColor
        optPurgeUndo.ForeColor = frmHighlight.ForeColor
        txtTargetTimeout.BackColor = frmHighlight.ForeColor
        txtTargetTimeout.ForeColor = frmHighlight.BackColor
        lblTargetTimeoutUnits.BackColor = frmHighlight.ForeColor
        lblTargetTimeoutUnits.ForeColor = frmHighlight.BackColor
        optTargetContinuous.Enabled = True
        optTargetPurgePauseRepeat.Enabled = True
        lblTargetTimeoutUnits.Enabled = True
        lblTargetPurgeUnits.Enabled = True
        lblTargetPauseUnits.Enabled = True
        txtTargetTimeout.Enabled = True
        txtTargetPurge.Enabled = True
        txtTargetPause.Enabled = True
        If optTargetContinuous.Value Then
            optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
            txtTargetPurge.BackColor = txtNotHighlight.BackColor
            txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
            txtTargetPause.BackColor = txtNotHighlight.BackColor
            txtTargetPause.ForeColor = txtNotHighlight.ForeColor
            lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
            lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
            optTargetContinuous.BackColor = frmHighlight.BackColor
            optTargetContinuous.ForeColor = frmHighlight.ForeColor
        ElseIf optTargetPurgePauseRepeat.Value Then
            optTargetContinuous.BackColor = frmNotHighlight.BackColor
            optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
            optTargetPurgePauseRepeat.BackColor = frmHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmHighlight.ForeColor
            txtTargetPurge.BackColor = frmHighlight.ForeColor
            txtTargetPurge.ForeColor = frmHighlight.BackColor
            txtTargetPause.BackColor = frmHighlight.ForeColor
            txtTargetPause.ForeColor = frmHighlight.BackColor
            lblTargetPurgeUnits.BackColor = frmHighlight.ForeColor
            lblTargetPurgeUnits.ForeColor = frmHighlight.BackColor
            lblTargetPauseUnits.BackColor = frmHighlight.ForeColor
            lblTargetPauseUnits.ForeColor = frmHighlight.BackColor
        End If
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
    Else
        optPurgeUndo.BackColor = frmNotHighlight.BackColor
        optPurgeUndo.ForeColor = frmNotHighlight.ForeColor
        txtTargetTimeout.BackColor = txtNotHighlight.BackColor
        txtTargetTimeout.ForeColor = txtNotHighlight.ForeColor
        lblTargetTimeoutUnits.BackColor = txtNotHighlight.BackColor
        lblTargetTimeoutUnits.ForeColor = txtNotHighlight.ForeColor
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
        lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
        lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeVolume_Click()
    If optPurgeVolume.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeAuxOnly.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        optPurgeVolume.BackColor = frmHighlight.BackColor
        optPurgeVolume.ForeColor = frmHighlight.ForeColor
        txtPurgeVolume.BackColor = frmHighlight.ForeColor
        txtPurgeVolume.ForeColor = frmHighlight.BackColor
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
        optTargetContinuous.Enabled = True
        optTargetPurgePauseRepeat.Enabled = True
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = True
        lblTargetPauseUnits.Enabled = True
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = True
        txtTargetPause.Enabled = True
    Else
        optPurgeVolume.BackColor = frmNotHighlight.BackColor
        optPurgeVolume.ForeColor = frmNotHighlight.ForeColor
        txtPurgeVolume.BackColor = txtNotHighlight.BackColor
        txtPurgeVolume.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeAuxOnly_Click()
    If optPurgeAuxOnly.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        chkPurgeAuxCan.Value = cYES
        chkUseAuxScale.Value = cYES
        chkPurgeCansInSeries.Value = cNO
        txtPurgeTime.text = "0"
        txtPurgeVolume.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        txtPurgeFlow.text = "0"
        optPurgeAuxOnly.BackColor = frmHighlight.BackColor
        optPurgeAuxOnly.ForeColor = frmHighlight.ForeColor
        txtPurgeAuxOnly.BackColor = frmHighlight.ForeColor
        txtPurgeAuxOnly.ForeColor = frmHighlight.BackColor
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
        optTargetContinuous.Enabled = False
        optTargetPurgePauseRepeat.Enabled = False
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = False
        lblTargetPauseUnits.Enabled = False
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = False
        txtTargetPause.Enabled = False
    Else
        optPurgeAuxOnly.BackColor = frmNotHighlight.BackColor
        optPurgeAuxOnly.ForeColor = frmNotHighlight.ForeColor
        txtPurgeAuxOnly.BackColor = txtNotHighlight.BackColor
        txtPurgeAuxOnly.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeProfile_Click()
    If optPurgeProfile.Value = cYES Then
'        txtPurgeTime = cNO
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeTime.text = "0"
        txtPurgeVolume.text = "0"
        txtPurgeAuxOnly.text = "0"
        txtPurgeWC.text = "0"
        txtPurgeTarget.text = "0"
        txtTargetTimeout.text = "0"
        optPurgeProfile.BackColor = frmHighlight.BackColor
        optPurgeProfile.ForeColor = frmHighlight.ForeColor
        txtPurgeProfile.BackColor = frmHighlight.ForeColor
        txtPurgeProfile.ForeColor = frmHighlight.BackColor
        lblPurgeFlow.Enabled = False
        txtPurgeFlow.Enabled = False
        optTargetContinuous.Enabled = False
        optTargetPurgePauseRepeat.Enabled = False
        lblTargetTimeoutUnits.Enabled = False
        lblTargetPurgeUnits.Enabled = False
        lblTargetPauseUnits.Enabled = False
        txtTargetTimeout.Enabled = False
        txtTargetPurge.Enabled = False
        txtTargetPause.Enabled = False
    Else
        optPurgeProfile.BackColor = frmNotHighlight.BackColor
        optPurgeProfile.ForeColor = frmNotHighlight.ForeColor
        txtPurgeProfile.BackColor = frmPurge.BackColor
        txtPurgeProfile.ForeColor = frmPurge.BackColor
    End If
End Sub

Private Sub optPurgeTarget_Click()
    If optPurgeTarget.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeWC.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeTime.text = "0"
        txtPurgeVolume.text = "0"
        txtPurgeAuxOnly.text = "0"
        txtPurgeWC.text = "0"
        optPurgeTarget.BackColor = frmHighlight.BackColor
        optPurgeTarget.ForeColor = frmHighlight.ForeColor
        txtPurgeTarget.BackColor = frmHighlight.ForeColor
        txtPurgeTarget.ForeColor = frmHighlight.BackColor
        txtTargetTimeout.BackColor = frmHighlight.ForeColor
        txtTargetTimeout.ForeColor = frmHighlight.BackColor
        lblTargetTimeoutUnits.BackColor = frmHighlight.ForeColor
        lblTargetTimeoutUnits.ForeColor = frmHighlight.BackColor
        optTargetContinuous.Enabled = True
        optTargetPurgePauseRepeat.Enabled = True
        lblTargetTimeoutUnits.Enabled = True
        lblTargetPurgeUnits.Enabled = True
        lblTargetPauseUnits.Enabled = True
        txtTargetTimeout.Enabled = True
        txtTargetPurge.Enabled = True
        txtTargetPause.Enabled = True
        If optTargetContinuous.Value Then
            optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
            txtTargetPurge.BackColor = txtNotHighlight.BackColor
            txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
            txtTargetPause.BackColor = txtNotHighlight.BackColor
            txtTargetPause.ForeColor = txtNotHighlight.ForeColor
            lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
            lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
            optTargetContinuous.BackColor = frmHighlight.BackColor
            optTargetContinuous.ForeColor = frmHighlight.ForeColor
        ElseIf optTargetPurgePauseRepeat.Value Then
            optTargetContinuous.BackColor = frmNotHighlight.BackColor
            optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
            optTargetPurgePauseRepeat.BackColor = frmHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmHighlight.ForeColor
            txtTargetPurge.BackColor = frmHighlight.ForeColor
            txtTargetPurge.ForeColor = frmHighlight.BackColor
            txtTargetPause.BackColor = frmHighlight.ForeColor
            txtTargetPause.ForeColor = frmHighlight.BackColor
            lblTargetPurgeUnits.BackColor = frmHighlight.ForeColor
            lblTargetPurgeUnits.ForeColor = frmHighlight.BackColor
            lblTargetPauseUnits.BackColor = frmHighlight.ForeColor
            lblTargetPauseUnits.ForeColor = frmHighlight.BackColor
        End If
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
    Else
        optPurgeTarget.BackColor = frmNotHighlight.BackColor
        optPurgeTarget.ForeColor = frmNotHighlight.ForeColor
        txtPurgeTarget.BackColor = txtNotHighlight.BackColor
        txtPurgeTarget.ForeColor = txtNotHighlight.ForeColor
        txtTargetTimeout.BackColor = txtNotHighlight.BackColor
        txtTargetTimeout.ForeColor = txtNotHighlight.ForeColor
        lblTargetTimeoutUnits.BackColor = txtNotHighlight.BackColor
        lblTargetTimeoutUnits.ForeColor = txtNotHighlight.ForeColor
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
        lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
        lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optPurgeWC_Click()
    If optPurgeWC.Value = cYES Then
        optNoPurge.Value = cNO
        optPurgeTime.Value = cNO
        optPurgeLiters.Value = cNO
        optPurgeVolume.Value = cNO
        optPurgeAuxOnly.Value = cNO
        optPurgeProfile.Value = cNO
        optPurgeTarget.Value = cNO
        optPurgeUndo.Value = cNO
        txtPurgeTime.text = "0"
        txtPurgeVolume.text = "0"
        txtPurgeAuxOnly.text = "0"
        txtPurgeTarget.text = "0"
        optPurgeWC.BackColor = frmHighlight.BackColor
        optPurgeWC.ForeColor = frmHighlight.ForeColor
        txtPurgeWC.BackColor = frmHighlight.ForeColor
        txtPurgeWC.ForeColor = frmHighlight.BackColor
        txtTargetTimeout.BackColor = frmHighlight.ForeColor
        txtTargetTimeout.ForeColor = frmHighlight.BackColor
        lblTargetTimeoutUnits.BackColor = frmHighlight.ForeColor
        lblTargetTimeoutUnits.ForeColor = frmHighlight.BackColor
        optTargetContinuous.Enabled = True
        optTargetPurgePauseRepeat.Enabled = True
        lblTargetTimeoutUnits.Enabled = True
        lblTargetPurgeUnits.Enabled = True
        lblTargetPauseUnits.Enabled = True
        txtTargetTimeout.Enabled = True
        txtTargetPurge.Enabled = True
        txtTargetPause.Enabled = True
        If optTargetContinuous.Value Then
            optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
            txtTargetPurge.BackColor = txtNotHighlight.BackColor
            txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
            txtTargetPause.BackColor = txtNotHighlight.BackColor
            txtTargetPause.ForeColor = txtNotHighlight.ForeColor
            lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
            lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
            optTargetContinuous.BackColor = frmHighlight.BackColor
            optTargetContinuous.ForeColor = frmHighlight.ForeColor
        ElseIf optTargetPurgePauseRepeat.Value Then
            optTargetContinuous.BackColor = frmNotHighlight.BackColor
            optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
            optTargetPurgePauseRepeat.BackColor = frmHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmHighlight.ForeColor
            txtTargetPurge.BackColor = frmHighlight.ForeColor
            txtTargetPurge.ForeColor = frmHighlight.BackColor
            txtTargetPause.BackColor = frmHighlight.ForeColor
            txtTargetPause.ForeColor = frmHighlight.BackColor
            lblTargetPurgeUnits.BackColor = frmHighlight.ForeColor
            lblTargetPurgeUnits.ForeColor = frmHighlight.BackColor
            lblTargetPauseUnits.BackColor = frmHighlight.ForeColor
            lblTargetPauseUnits.ForeColor = frmHighlight.BackColor
        End If
        lblPurgeFlow.Enabled = True
        txtPurgeFlow.Enabled = True
    Else
        optPurgeWC.BackColor = frmNotHighlight.BackColor
        optPurgeWC.ForeColor = frmNotHighlight.ForeColor
        txtPurgeWC.BackColor = txtNotHighlight.BackColor
        txtPurgeWC.ForeColor = txtNotHighlight.ForeColor
        txtTargetTimeout.BackColor = txtNotHighlight.BackColor
        txtTargetTimeout.ForeColor = txtNotHighlight.ForeColor
        lblTargetTimeoutUnits.BackColor = txtNotHighlight.BackColor
        lblTargetTimeoutUnits.ForeColor = txtNotHighlight.ForeColor
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
        lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
        lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optStartAfter_Click()
    If optStartAfter.Value = cYES Then
        txtStartAtDate.text = Format(Now(), "DD/MM/YYYY hh:mm")
        optStartNow.Value = cNO
        optStartAt.Value = cNO
        optStartAfter.BackColor = frmHighlight.BackColor
        optStartAfter.ForeColor = frmHighlight.ForeColor
        txtStartAfterMin.BackColor = frmHighlight.ForeColor
        txtStartAfterMin.ForeColor = frmHighlight.BackColor
    Else
        optStartAfter.BackColor = frmNotHighlight.BackColor
        optStartAfter.ForeColor = frmNotHighlight.ForeColor
        txtStartAfterMin.BackColor = txtNotHighlight.BackColor
        txtStartAfterMin.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optStartAt_Click()
    If optStartAt.Value = cYES Then
        txtStartAfterMin.text = "0"
        optStartAfter.Value = cNO
        optStartNow.Value = cNO
        optStartAt.BackColor = frmHighlight.BackColor
        optStartAt.ForeColor = frmHighlight.ForeColor
        txtStartAtDate.BackColor = frmHighlight.ForeColor
        txtStartAtDate.ForeColor = frmHighlight.BackColor
    Else
        optStartAt.BackColor = frmNotHighlight.BackColor
        optStartAt.ForeColor = frmNotHighlight.ForeColor
        txtStartAtDate.BackColor = txtNotHighlight.BackColor
        txtStartAtDate.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optStartNow_Click()
    If optStartNow.Value = cYES Then
        txtStartAfterMin.text = "0"
        txtStartAtDate.text = Format(Now(), "DD/MM/YYYY hh:mm")
        optStartAfter.Value = cNO
        optStartAt.Value = cNO
        optStartNow.BackColor = frmHighlight.BackColor
        optStartNow.ForeColor = frmHighlight.ForeColor
    Else
        optStartNow.BackColor = frmNotHighlight.BackColor
        optStartNow.ForeColor = frmNotHighlight.ForeColor
    End If
End Sub

Private Sub optTargetContinuous_Click()
    If optPurgeTarget.Value = cYES Then
        If optTargetContinuous.Value = True Then
            optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
            txtTargetPurge.BackColor = txtNotHighlight.BackColor
            txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
            txtTargetPause.BackColor = txtNotHighlight.BackColor
            txtTargetPause.ForeColor = txtNotHighlight.ForeColor
            lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
            lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
            optTargetContinuous.BackColor = frmHighlight.BackColor
            optTargetContinuous.ForeColor = frmHighlight.ForeColor
        Else
            optTargetContinuous.BackColor = frmNotHighlight.BackColor
            optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        End If
    Else
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
        lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
        lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optTargetPurgePauseRepeat_Click()
    If optPurgeTarget.Value = cYES Then
        If optTargetPurgePauseRepeat.Value Then
            optTargetContinuous.BackColor = frmNotHighlight.BackColor
            optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
            optTargetPurgePauseRepeat.BackColor = frmHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmHighlight.ForeColor
            txtTargetPurge.BackColor = frmHighlight.ForeColor
            txtTargetPurge.ForeColor = frmHighlight.BackColor
            txtTargetPause.BackColor = frmHighlight.ForeColor
            txtTargetPause.ForeColor = frmHighlight.BackColor
            lblTargetPurgeUnits.BackColor = frmHighlight.ForeColor
            lblTargetPurgeUnits.ForeColor = frmHighlight.BackColor
            lblTargetPauseUnits.BackColor = frmHighlight.ForeColor
            lblTargetPauseUnits.ForeColor = frmHighlight.BackColor
        Else
            optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
            optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
            txtTargetPurge.BackColor = txtNotHighlight.BackColor
            txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
            txtTargetPause.BackColor = txtNotHighlight.BackColor
            txtTargetPause.ForeColor = txtNotHighlight.ForeColor
            lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
            lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
            lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
        End If
    Else
        optTargetContinuous.BackColor = frmNotHighlight.BackColor
        optTargetContinuous.ForeColor = frmNotHighlight.ForeColor
        optTargetPurgePauseRepeat.BackColor = frmNotHighlight.BackColor
        optTargetPurgePauseRepeat.ForeColor = frmNotHighlight.ForeColor
        txtTargetPurge.BackColor = txtNotHighlight.BackColor
        txtTargetPurge.ForeColor = txtNotHighlight.ForeColor
        txtTargetPause.BackColor = txtNotHighlight.BackColor
        txtTargetPause.ForeColor = txtNotHighlight.ForeColor
        lblTargetPurgeUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPurgeUnits.ForeColor = txtNotHighlight.ForeColor
        lblTargetPauseUnits.BackColor = txtNotHighlight.BackColor
        lblTargetPauseUnits.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub optWCM_Click()
    If optWcm.Value = cYES Then
        txtLoadTime.text = "0"
        txtTargetWt.text = "0"
        txtLoadBreakthrough.text = "0"
        txtFIDmg.text = "0"
        txtLoadRate.text = "15"
        optLoadTime.Value = 0
        optNoLoad.Value = 0
        optLoadweight.Value = 0
        optLoadBreakthrough.Value = 0
        optFIDBreakthrough.Value = 0
        optWcm.BackColor = frmHighlight.BackColor
        optWcm.ForeColor = frmHighlight.ForeColor
        txtWorkCapMult.BackColor = frmHighlight.ForeColor
        txtWorkCapMult.ForeColor = frmHighlight.BackColor
        txtEPAFill.BackColor = frmHighlight.ForeColor
        txtEPAFill.ForeColor = frmHighlight.BackColor
    Else
        optWcm.BackColor = frmNotHighlight.BackColor
        optWcm.ForeColor = frmNotHighlight.ForeColor
        txtWorkCapMult.BackColor = txtNotHighlight.BackColor
        txtWorkCapMult.ForeColor = txtNotHighlight.ForeColor
        txtEPAFill.BackColor = txtNotHighlight.BackColor
        txtEPAFill.ForeColor = txtNotHighlight.ForeColor
    End If
End Sub

Private Sub pnlDispRcpNum_Click()
    ' select a different recipe from the master list
    frmSearchRcp.Show
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tmrUpdate_Timer()
    pnlDispRcpNum.Refresh
    tmrUpdate.Enabled = False
End Sub

Private Sub txtADF_HeaterSP_Change()
    txtADF_HeaterSP.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtAuxScaleNo_Change()
    chkUseAuxScale.BackColor = frmNotHighlight.BackColor
    txtAuxScaleNo.BackColor = txtNotHighlight.BackColor
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        txtAuxScaleNo.text = Format(DspRecipe.AuxScaleNo, "#0")
    End If
End Sub

Private Sub txtButnPercent_Change()
    txtButnPercent.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPurgeLiters_Change()
    txtPurgeLiters.BackColor = txtNotHighlight.BackColor
    If IsNumeric(txtPurgeLiters) And optPurgeLiters = cYES Then PurgeTime
End Sub

Private Sub txtPurgeOvenSP_Change()
    txtPurgeOvenSP.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPurgeTime_Change()
    txtPurgeTime.BackColor = txtNotHighlight.BackColor
    If IsNumeric(txtPurgeTime) And optPurgeTime = 1 Then PurgeVolume
End Sub

Private Sub txtPurgeV_Change()
    txtPurgeV.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPurgeVolume_Change()
    txtPurgeVolume.BackColor = txtNotHighlight.BackColor
    If IsNumeric(txtPurgeVolume) And optPurgeVolume = cYES Then PurgeTime
End Sub

Private Sub txtEPAFill_Change()
    txtEPAFill.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtFIDmg_Change()
    txtFIDmg.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtIDLoad_Change()
    txtIDLoad.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtIDPurge_Change()
    txtIDPurge.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtIDVent_Change()
    txtIDVent.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtLiveFuelChgFreq_Change()
    txtLiveFuelChgFreq.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtloadbreakthrough_Change()
    txtLoadBreakthrough.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtLoadL_Change()
    txtLoadL.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtLoadRate_Change()
    txtLoadRate.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtLoadTime_Change()
    txtLoadTime.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtMaxLoadTime_Change()
    txtMaxLoadTime.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtNitrogenFlow_Change()
    txtNitrogenFlow.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPauseLeakTime_Change()
    txtPauseLeakTime.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPauseLoadTime_Change()
    txtPauseLoadTime.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPausePurgeTime_Change()
    txtPausePurgeTime.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPFCycle_Change()
    txtPFCycle.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtPrimaryScaleNo_Change()
    chkPrimaryScale.BackColor = frmNotHighlight.BackColor
    txtPrimaryScaleNo.BackColor = txtNotHighlight.BackColor
    If RecipeMode = STATIONMODE And USINGHARDPIPEDSCALES Then
        txtPrimaryScaleNo.text = Format(DspRecipe.PriScaleNo, "#0")
    End If
End Sub

Private Sub txtPurgeFlow_Change()
    txtPurgeFlow.BackColor = txtNotHighlight.BackColor
    If IsNumeric(txtPurgeFlow) Then
        If IsNumeric(txtPurgeTime) And optPurgeTime = cYES Then PurgeVolume
        If IsNumeric(txtPurgeVolume) And optPurgeVolume = cYES Then PurgeTime
    End If
End Sub
Private Sub txtPurgeL_Change()
    txtPurgeL.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtStartAfterMin_Change()
    txtStartAfterMin.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtStartAtDate_Change()
    txtStartAtDate.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtTargetWt_Change()
    txtTargetWt.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtVentL_Change()
    txtVentL.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtVentV_Change()
    txtVentV.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub txtWorkCapmult_Change()
    txtWorkCapMult.BackColor = txtNotHighlight.BackColor
End Sub

Private Sub Print_AllMasters()
' Procedure Name:   Print_AllMasters
' Created By:       Brunrose
' Description:
'
Dim Idx As Integer
Dim oldFont As New StdFont
Dim dB As DAO.Database
Dim rS As DAO.Recordset
Dim sPath As String
Dim rsCrit As String
Dim curRcp As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 222

    ' open master recipes datatable
    sPath = FILEPATH_rcp & DATARCP
    Set dB = DBEngine.OpenDatabase(sPath)
    rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Number] ASC"
    Set rS = dB.OpenRecordset(rsCrit, dbOpenDynaset)

    ' number of recipes
    rS.GetRows
    rS.MoveFirst
    If rS.RecordCount = 0 Then
        ' Save current printer font
        oldFont = Printer.Font
        Printer.Font = REPORTFONT
        ' Print blank line(s)
        Print_Line ""
        Print_Line ""
        ' No Data
        Print_Center "No defined master recipes"
        Printer.EndDoc
        Printer.Font = oldFont
    Else
        ' save current Recipe#
        curRcp = DispRcp
        For Idx = 1 To rS.RecordCount
            ' print one recipe
            frmRecipe.LoadNewRcp CInt(rS("Number"))
            lblMessage.Caption = ""
            Set pbCapture.Picture = CaptureForm(Me)
            PrintPictureToFitPage Printer, pbCapture.Picture
            Printer.EndDoc
            Set pbCapture.Picture = Nothing
            rS.MoveNext
        Next Idx
        ' restore current Recipe#
        DispRcp = curRcp
        frmRecipe.LoadNewRcp CInt(DispRcp)
        rS.MoveFirst
    End If
    'DONE
    rS.Close
    dB.Close
    
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

Private Sub Print_AllStationRecipes()
' Procedure Name:   Print_AllStationRecipes
' Created By:       Brunrose
' Description:
'
Dim iStn As Integer
Dim iShift As Integer

If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 90, 22772

    ' cycle thru all stations and shifts
    For iStn = 1 To LAST_STN
        For iShift = 1 To NR_SHIFT
            ' open the recipe
            GetRecipe STATIONMODE, iStn, iShift
            DspRcpToScreen
            Chgs = False
            txtRecipeName.text = txtRecipeName.text & "  ---  Station " & Format(iStn, "#0") & "  Shift " & Format(iShift, "0")
            lblMessage.Caption = ""
            ' capture the recipe
            Set pbCapture.Picture = CaptureForm(Me)
            ' print the recipe
            PrintPictureToFitPage Printer, pbCapture.Picture
            Printer.EndDoc
            ' short delay
            DoEvents
'            Delay_Box "", PAUSEDELAY, msgNOSHOW
        Next iShift
    Next iStn
    ' clear capture box
    Set pbCapture.Picture = Nothing
    DoEvents
    ' restore current Stn/Shift Recipe
    GetRecipe STATIONMODE, DispStn, DispShift
    DspRcpToScreen
    Chgs = False
    lblMessage.Caption = ""
    
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

Private Sub SaveRemoteMasterRecipe(ByVal iRcp As Integer)
'
'        Save Remote Master Recipe Information Record
'
'        frmRemoteRcp.Hide
        frmRemoteRcp.Show
        ' Open existing Remote Master Recipe Information Record (if any)
        frmRemoteRcp.adoRemoteRecipes.RecordSource = "SELECT * FROM [MasterRecipe] WHERE [MasterRecipe].[Number] = " & iRcp & " "
'        Set .Fields = dbDbase.OpenRecordset(Criteria, dbOpenDynaset)
        frmRemoteRcp.adoRemoteRecipes.Refresh
        frmRemoteRcp.dbgRemoteRecipes.Refresh
        
        With frmRemoteRcp.adoRemoteRecipes.Recordset
        
            If .BOF Then
                .AddNew
                .Fields("Number").Value = iRcp
            Else
              .MoveLast
              .MoveFirst
            End If
               
            If .RecordCount = 1 Then
                ' Update Remote Master Recipe Information Record
                .Fields("Name").Value = DspRecipe.Name
                
                .Fields("CycleType").Value = DspRecipe.CycleType
                .Fields("CycleTypeDesc").Value = CycleTypeDesc(DspRecipe.CycleType)
                        
                .Fields("Load_Method").Value = DspRecipe.Load_Method
                .Fields("Load_MethodDesc").Value = LoadMethodDesc(DspRecipe.Load_Method)
                .Fields("NitrogenFlow").Value = DspRecipe.NitrogenFlow
                .Fields("Load_Rate").Value = DspRecipe.Load_Rate
                .Fields("Mix_Percent").Value = DspRecipe.Mix_Percent
                .Fields("WC_Mult").Value = DspRecipe.WC_Mult
                .Fields("EPAFill").Value = DspRecipe.EPAFill
                .Fields("Load_Wt").Value = DspRecipe.Load_Wt
                .Fields("LoadBreakthrough").Value = DspRecipe.LoadBreakthrough
                .Fields("Load_Time").Value = DspRecipe.Load_Time
                .Fields("Purge_Method").Value = DspRecipe.Purge_Method
                .Fields("Purge_MethodDesc").Value = PurgeMethodDesc(DspRecipe.Purge_Method)
                .Fields("Purge_AuxTime").Value = DspRecipe.Purge_AuxTime
                .Fields("Purge_Time").Value = DspRecipe.Purge_Time
                .Fields("Purge_Flow").Value = DspRecipe.Purge_Flow
                .Fields("Purge_Liters").Value = DspRecipe.Purge_Liters
                .Fields("Purge_Can_Vol").Value = DspRecipe.Purge_Can_Vol
                .Fields("Purge_ProfileNumber").Value = DspRecipe.Purge_ProfileNumber
                .Fields("Purge_TargetMode").Value = DspRecipe.Purge_TargetMode
                .Fields("Purge_TargetModeDesc").Value = PurgeTargetDesc(DspRecipe.Purge_TargetMode)
                .Fields("Purge_TargetWC").Value = DspRecipe.Purge_TargetWC
                .Fields("Purge_TargetWeight").Value = DspRecipe.Purge_TargetWeight
                .Fields("Purge_MaxVolumes").Value = DspRecipe.Purge_MaxVolumes
                .Fields("Purge_TargetPurge").Value = DspRecipe.Purge_TargetPurge
                .Fields("Purge_TargetPause").Value = DspRecipe.Purge_TargetPause
                
                .Fields("PurgeAuxCan").Value = DspRecipe.PurgeAuxCan
                .Fields("PurgeCansInSeries").Value = DspRecipe.PurgeCansInSeries
                .Fields("PurgeInOven").Value = DspRecipe.PurgeOven
                .Fields("PurgeOvenSP").Value = DspRecipe.PurgeOvenSP
                .Fields("UseAuxScale").Value = DspRecipe.UseAuxScale
                .Fields("AuxScaleNo").Value = DspRecipe.AuxScaleNo
                .Fields("PauseLeakTime").Value = DspRecipe.PauseLeakTime
                .Fields("PauseLoadTime").Value = DspRecipe.PauseLoadTime
                .Fields("PausePurgeTime").Value = DspRecipe.PausePurgeTime
                .Fields("UsePriScale").Value = DspRecipe.UsePriScale
                .Fields("PriScaleNo").Value = DspRecipe.PriScaleNo
                .Fields("PauseAfterLeak").Value = DspRecipe.PauseAfterLeak
                .Fields("PauseAfterLoad").Value = DspRecipe.PauseAfterLoad
                .Fields("PauseAfterLoadForOper").Value = DspRecipe.PauseAfterLoadForOper
                .Fields("PauseAfterPurge").Value = DspRecipe.PauseAfterPurge
                .Fields("PauseAfterPurgeForOper").Value = DspRecipe.PauseAfterPurgeForOper
'                .Fields("TargetConcentration").Value = DspRecipe.TargetConcentration
'                .Fields("DwellTime").Value = DspRecipe.DwellTime
                .Fields("LeakCheck").Value = DspRecipe.LeakCheck
                .Fields("LeakPrimary").Value = DspRecipe.LeakPrimary
                .Fields("LeakAux").Value = DspRecipe.LeakAux
'                .Fields("UseAnalyzer").Value = DspRecipe.UseAnalyzer
                .Fields("MaxLoadTime").Value = DspRecipe.MaxLoadTime
                .Fields("UseHiRangeMFC").Value = DspRecipe.UseHiRangeMFC
                .Fields("UseLoadRatePID").Value = DspRecipe.UseLoadRatePID
                
                .Fields("IDLoad").Value = DspRecipe.IDLoad
                .Fields("LoadL").Value = DspRecipe.LoadL
                .Fields("LoadV").Value = DspRecipe.LoadV
                .Fields("IDPurge").Value = DspRecipe.IDPurge
                .Fields("PurgeL").Value = DspRecipe.PurgeL
                .Fields("PurgeV").Value = DspRecipe.PurgeV
                .Fields("IDVent").Value = DspRecipe.IDVent
                .Fields("VentL").Value = DspRecipe.VentL
                .Fields("VentV").Value = DspRecipe.VentV
                
                .Fields("LiveFuel").Value = DspRecipe.LiveFuel
                .Fields("LiveFuelChgAuto").Value = DspRecipe.LiveFuelChgAuto
                .Fields("LiveFuelChgFreq").Value = DspRecipe.LiveFuelChgFreq
                .Fields("ADF_Heater").Value = DspRecipe.ADF_Heater
                .Fields("ADF_HeaterSP").Value = DspRecipe.ADF_HeaterSP
                
                ' start method
                .Fields("StartMethod").Value = DspRecipe.StartMethod
                .Fields("StartMethodDesc").Value = StartMethodDesc(DspRecipe.StartMethod)
                .Fields("StartDelay").Value = DspRecipe.StartDelay
                .Fields("StartDate").Value = DspRecipe.StartDate
                    
                ' end method
                .Fields("EndMethod").Value = DspRecipe.EndMethod
                .Fields("EndMethodDesc").Value = EndMethodDesc(DspRecipe.EndMethod)
                .Fields("EndMaximumCycles").Value = DspRecipe.EndMaximumCycles
                .Fields("EndMinimumCycles").Value = DspRecipe.EndMinimumCycles
                .Fields("EndConsecutiveCycles").Value = DspRecipe.EndConsecutiveCycles
                .Fields("EndWeightTolerance").Value = DspRecipe.EndWeightTolerance
                .Fields("UpdateCanWc").Value = DspRecipe.UpdateCanWc
                .Fields("Cycles").Value = DspRecipe.Cycles
                    
                ' aux outputs
                .Fields("AuxOutputs").Value = DspRecipe.AuxOutputs
                .Fields("AuxOutput1_Load").Value = DspRecipe.AuxOutputs_Load(1)
                .Fields("AuxOutput1_Purge").Value = DspRecipe.AuxOutputs_Purge(1)
                .Fields("AuxOutput2_Load").Value = DspRecipe.AuxOutputs_Load(2)
                .Fields("AuxOutput2_Purge").Value = DspRecipe.AuxOutputs_Purge(2)
                .Fields("AuxOutput3_Load").Value = DspRecipe.AuxOutputs_Load(3)
                .Fields("AuxOutput3_Purge").Value = DspRecipe.AuxOutputs_Purge(3)
                .Fields("AuxOutput4_Load").Value = DspRecipe.AuxOutputs_Load(4)
                .Fields("AuxOutput4_Purge").Value = DspRecipe.AuxOutputs_Purge(4)
                .Update
            Else
                Write_ELog "RemoteRcp Update Failure - Multiple Records Returned for Rcp# " & Format(iRcp, "#,##0")
            End If
            
        End With
        
        ' reset RemoteRecipes RecordSource
        frmRemoteRcp.adoRemoteRecipes.RecordSource = "SELECT * FROM [MasterRecipe] ORDER BY [MasterRecipe].[Number] ASC"
        frmRemoteRcp.adoRemoteRecipes.Refresh
        frmRemoteRcp.dbgRemoteRecipes.Refresh
                    

        Unload frmRemoteRcp

End Sub


