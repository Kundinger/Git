VERSION 5.00
Begin VB.Form frmWizardPrg 
   BackColor       =   &H00000000&
   Caption         =   "Purge Wizard"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6930
   Icon            =   "frmWizardPrg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6930
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9330
      Begin VB.Frame frmPrefill 
         Caption         =   "Recipe Purge Parameters"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   1680
         Left            =   120
         TabIndex        =   2
         Top             =   2520
         Width           =   6705
         Begin VB.CommandButton cmdApply 
            Caption         =   "Apply"
            DisabledPicture =   "frmWizardPrg.frx":57E2
            DownPicture     =   "frmWizardPrg.frx":6424
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   6000
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmWizardPrg.frx":7066
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            UseMaskColor    =   -1  'True
            Width           =   600
         End
         Begin VB.TextBox txtPurgeFlow 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   2025
            TabIndex        =   4
            Text            =   "28"
            ToolTipText     =   "Sample Gas Pre-Fill time in seconds"
            Top             =   600
            Width           =   900
         End
         Begin VB.TextBox txtPurgeTime 
            Alignment       =   2  'Center
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
            Height          =   285
            Left            =   2025
            TabIndex        =   3
            Text            =   "22"
            ToolTipText     =   "Nitrogen Pre-Fill time in seconds"
            Top             =   300
            Width           =   900
         End
         Begin VB.Label Label3 
            Caption         =   "Purge Flow Rate"
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
            Left            =   240
            TabIndex        =   8
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "min"
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
            Left            =   2955
            TabIndex        =   7
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label5 
            Caption         =   "Purge Time"
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
            Left            =   240
            TabIndex        =   6
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "slpm"
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
            Left            =   2955
            TabIndex        =   5
            Top             =   630
            Width           =   390
         End
      End
      Begin VB.Frame frmRecipe 
         Caption         =   "Purge Calculations"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   2295
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6705
         Begin VB.TextBox txtPurgeLiters 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4545
            TabIndex        =   14
            Text            =   "0.025"
            ToolTipText     =   "flow rate in slpm"
            Top             =   840
            Width           =   795
         End
         Begin VB.TextBox txtPurgeDuration 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   4440
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "3223"
            Top             =   360
            Width           =   900
         End
         Begin VB.Label lblPurgeLitersUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "liters"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   5520
            TabIndex        =   13
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lblPurgeLiters 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Number Of Liters  to Purge"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   3105
         End
         Begin VB.Label lblPurgeDuration 
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Purge Duration"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   120
            TabIndex        =   11
            Top             =   375
            Width           =   2400
         End
         Begin VB.Label lblPurgeDurationUnits 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   240
            Left            =   5520
            TabIndex        =   10
            Top             =   375
            Width           =   995
         End
      End
   End
End
Attribute VB_Name = "frmWizardPrg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  Module frmWizardPrg
Option Explicit
Private InitialLoadComplete As Boolean

Private Sub cmdApply_Click()
'    frmRecipe.Show
    PurgeWizardUpdater ValueFromText(txtPurgeTime.text), ValueFromText(txtPurgeFlow.text)
'    frmRecipe.Refresh
    Unload frmWizardPrg
End Sub

Private Sub Form_Load()
Dim netliters As Single
    InitialLoadComplete = False
    
    txtPurgeTime.text = Format(StationRecipe(DispStn, DispShift).Purge_Time, "###,##0")
    txtPurgeFlow.text = Format(StationRecipe(DispStn, DispShift).Purge_Flow, "###,##0.0##")
    
    txtPurgeDuration.text = txtPurgeTime.text
    netliters = ValueFromText(txtPurgeFlow.text) * ValueFromText(txtPurgeTime.text)
    txtPurgeLiters.text = Format(netliters, "##,##0.#")
    
    InitialLoadComplete = True
End Sub

Private Sub txtPurgeDuration_Change()
    If IsNumeric(txtPurgeDuration.text) Then
        If InitialLoadComplete Then UpdateProjections
   End If
End Sub

Private Sub txtPurgeLiters_Change()
    If IsNumeric(txtPurgeLiters.text) Then
        If InitialLoadComplete Then UpdateProjections
   End If
End Sub

Private Sub UpdateProjections()
Dim CalcFlowRate As Single
Dim InjSeconds As Single
Dim InjMass As Single
Dim InjLiters As Single
    If IsNumeric(txtPurgeDuration.text) Then
        txtPurgeTime.text = txtPurgeDuration.text
    End If
    If IsNumeric(txtPurgeDuration.text) And IsNumeric(txtPurgeLiters.text) Then
        CalcFlowRate = (ValueFromText(txtPurgeLiters.text) / ValueFromText(txtPurgeDuration.text))
        txtPurgeFlow.text = Format(CalcFlowRate, "####0.0##")
    End If
End Sub

