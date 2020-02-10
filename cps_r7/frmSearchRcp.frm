VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form frmSearchRcp 
   BackColor       =   &H80000005&
   Caption         =   "Master Recipes"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "frmSearchRcp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   14985
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dgRecipes 
      Bindings        =   "frmSearchRcp.frx":57E2
      Height          =   9255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   16325
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Recipes"
      ColumnCount     =   79
      BeginProperty Column00 
         DataField       =   "Number"
         Caption         =   "Number"
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
         DataField       =   "Name"
         Caption         =   "Name"
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
      BeginProperty Column03 
         DataField       =   "Load_MethodDesc"
         Caption         =   "Load_MethodDesc"
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
         DataField       =   "NitrogenFlow"
         Caption         =   "NitrogenFlow"
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
         DataField       =   "Load_Rate"
         Caption         =   "Load_Rate"
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
         DataField       =   "UseHiRangeMFC"
         Caption         =   "UseHiRangeMFC"
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
         DataField       =   "Mix_Percent"
         Caption         =   "Mix_Percent"
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
         DataField       =   "WC_Mult"
         Caption         =   "WC_Mult"
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
         DataField       =   "EPAFill"
         Caption         =   "EPAFill"
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
         DataField       =   "Load_Wt"
         Caption         =   "Load_Wt"
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
         DataField       =   "LoadBreakthrough"
         Caption         =   "LoadBreakthrough"
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
      BeginProperty Column12 
         DataField       =   "FidBreakthrough"
         Caption         =   "FidBreakthrough"
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
      BeginProperty Column13 
         DataField       =   "FIDmg"
         Caption         =   "FIDmg"
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
      BeginProperty Column14 
         DataField       =   "Load_Time"
         Caption         =   "Load_Time"
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
      BeginProperty Column15 
         DataField       =   "Purge_MethodDesc"
         Caption         =   "Purge_MethodDesc"
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
      BeginProperty Column16 
         DataField       =   "Purge_Flow"
         Caption         =   "Purge_Flow"
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
      BeginProperty Column17 
         DataField       =   "Purge_Can_Vol"
         Caption         =   "Purge_Can_Vol"
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
      BeginProperty Column18 
         DataField       =   "Purge_Time"
         Caption         =   "Purge_Time"
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
      BeginProperty Column19 
         DataField       =   "Purge_AuxTime"
         Caption         =   "Purge_AuxTime"
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
      BeginProperty Column20 
         DataField       =   "Purge_ProfileNumber"
         Caption         =   "Purge_ProfileNumber"
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
      BeginProperty Column21 
         DataField       =   "Purge_TargetModeDesc"
         Caption         =   "Purge_TargetModeDesc"
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
      BeginProperty Column22 
         DataField       =   "Purge_TargetWeight"
         Caption         =   "Purge_TargetWeight"
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
      BeginProperty Column23 
         DataField       =   "Purge_MaxVolumes"
         Caption         =   "Purge_MaxVolumes"
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
      BeginProperty Column24 
         DataField       =   "Purge_TargetPurge"
         Caption         =   "Purge_TargetPurge"
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
      BeginProperty Column25 
         DataField       =   "Purge_TargetPause"
         Caption         =   "Purge_TargetPause"
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
      BeginProperty Column26 
         DataField       =   "UseAuxScale"
         Caption         =   "UseAuxScale"
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
      BeginProperty Column27 
         DataField       =   "PurgeAuxCan"
         Caption         =   "PurgeAuxCan"
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
      BeginProperty Column28 
         DataField       =   "AuxScaleNo"
         Caption         =   "AuxScaleNo"
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
      BeginProperty Column29 
         DataField       =   "PauseAfterLeak"
         Caption         =   "PauseAfterLeak"
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
      BeginProperty Column30 
         DataField       =   "PauseLeakTime"
         Caption         =   "PauseLeakTime"
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
      BeginProperty Column31 
         DataField       =   "PauseAfterLoad"
         Caption         =   "PauseAfterLoad"
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
      BeginProperty Column32 
         DataField       =   "PauseLoadTime"
         Caption         =   "PauseLoadTime"
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
      BeginProperty Column33 
         DataField       =   "PauseAfterPurge"
         Caption         =   "PauseAfterPurge"
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
      BeginProperty Column34 
         DataField       =   "PausePurgeTime"
         Caption         =   "PausePurgeTime"
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
      BeginProperty Column35 
         DataField       =   "PrimaryScale"
         Caption         =   "PrimaryScale"
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
      BeginProperty Column36 
         DataField       =   "PrimaryScaleNo"
         Caption         =   "PrimaryScaleNo"
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
      BeginProperty Column37 
         DataField       =   "TargetConcentration"
         Caption         =   "TargetConcentration"
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
      BeginProperty Column38 
         DataField       =   "DwellTime"
         Caption         =   "DwellTime"
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
      BeginProperty Column39 
         DataField       =   "LeakCheck"
         Caption         =   "LeakCheck"
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
      BeginProperty Column40 
         DataField       =   "LeakPrimary"
         Caption         =   "LeakPrimary"
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
      BeginProperty Column41 
         DataField       =   "LeakAux"
         Caption         =   "LeakAux"
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
      BeginProperty Column42 
         DataField       =   "UseAnalyzer"
         Caption         =   "UseAnalyzer"
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
      BeginProperty Column43 
         DataField       =   "MaxLoadTime"
         Caption         =   "MaxLoadTime"
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
      BeginProperty Column44 
         DataField       =   "IDLoad"
         Caption         =   "IDLoad"
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
      BeginProperty Column45 
         DataField       =   "LoadL"
         Caption         =   "LoadL"
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
      BeginProperty Column46 
         DataField       =   "IDPurge"
         Caption         =   "IDPurge"
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
      BeginProperty Column47 
         DataField       =   "PurgeL"
         Caption         =   "PurgeL"
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
      BeginProperty Column48 
         DataField       =   "IDVent"
         Caption         =   "IDVent"
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
      BeginProperty Column49 
         DataField       =   "VentL"
         Caption         =   "VentL"
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
      BeginProperty Column50 
         DataField       =   "LoadV"
         Caption         =   "LoadV"
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
      BeginProperty Column51 
         DataField       =   "PurgeV"
         Caption         =   "PurgeV"
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
      BeginProperty Column52 
         DataField       =   "VentV"
         Caption         =   "VentV"
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
      BeginProperty Column53 
         DataField       =   "LiveFuel"
         Caption         =   "LiveFuel"
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
      BeginProperty Column54 
         DataField       =   "LiveFuelChgAuto"
         Caption         =   "LiveFuelChgAuto"
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
      BeginProperty Column55 
         DataField       =   "LiveFuelChgFreq"
         Caption         =   "LiveFuelChgFreq"
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
      BeginProperty Column56 
         DataField       =   "ADF_Heater"
         Caption         =   "ADF_Heater"
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
      BeginProperty Column57 
         DataField       =   "ADF_HeaterSP"
         Caption         =   "ADF_HeaterSP"
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
      BeginProperty Column58 
         DataField       =   "StartMethodDesc"
         Caption         =   "StartMethodDesc"
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
      BeginProperty Column59 
         DataField       =   "StartDelay"
         Caption         =   "StartDelay"
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
      BeginProperty Column60 
         DataField       =   "StartDate"
         Caption         =   "StartDate"
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
      BeginProperty Column61 
         DataField       =   "AuxOutputs"
         Caption         =   "AuxOutputs"
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
      BeginProperty Column62 
         DataField       =   "AuxOutput1_Load"
         Caption         =   "AuxOutput1_Load"
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
      BeginProperty Column63 
         DataField       =   "AuxOutput2_Load"
         Caption         =   "AuxOutput2_Load"
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
      BeginProperty Column64 
         DataField       =   "AuxOutput3_Load"
         Caption         =   "AuxOutput3_Load"
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
      BeginProperty Column65 
         DataField       =   "AuxOutput4_Load"
         Caption         =   "AuxOutput4_Load"
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
      BeginProperty Column66 
         DataField       =   "AuxOutput1_Purge"
         Caption         =   "AuxOutput1_Purge"
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
      BeginProperty Column67 
         DataField       =   "AuxOutput2_Purge"
         Caption         =   "AuxOutput2_Purge"
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
      BeginProperty Column68 
         DataField       =   "AuxOutput3_Purge"
         Caption         =   "AuxOutput3_Purge"
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
      BeginProperty Column69 
         DataField       =   "AuxOutput4_Purge"
         Caption         =   "AuxOutput4_Purge"
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
      BeginProperty Column70 
         DataField       =   "EndMethodDesc"
         Caption         =   "EndMethodDesc"
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
      BeginProperty Column71 
         DataField       =   "EndWeightTolerance"
         Caption         =   "EndWeightTolerance"
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
      BeginProperty Column72 
         DataField       =   "EndConsecutiveCycles"
         Caption         =   "EndConsecutiveCycles"
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
      BeginProperty Column73 
         DataField       =   "EndMinimumCycles"
         Caption         =   "EndMinimumCycles"
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
      BeginProperty Column74 
         DataField       =   "Load_Method"
         Caption         =   "Load_Method"
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
      BeginProperty Column75 
         DataField       =   "Purge_Method"
         Caption         =   "Purge_Method"
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
      BeginProperty Column76 
         DataField       =   "Purge_TargetMode"
         Caption         =   "Purge_TargetMode"
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
      BeginProperty Column77 
         DataField       =   "StartMethod"
         Caption         =   "StartMethod"
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
      BeginProperty Column78 
         DataField       =   "EndMethod"
         Caption         =   "EndMethod"
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
         BeginProperty Column00 
            ColumnWidth     =   629.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1319.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1395.213
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   1184.882
         EndProperty
         BeginProperty Column18 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   1574.929
         EndProperty
         BeginProperty Column21 
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column22 
            ColumnWidth     =   1560.189
         EndProperty
         BeginProperty Column23 
            ColumnWidth     =   1484.787
         EndProperty
         BeginProperty Column24 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column25 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column26 
            ColumnWidth     =   1035.213
         EndProperty
         BeginProperty Column27 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column28 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column29 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column30 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column31 
            ColumnWidth     =   1214.929
         EndProperty
         BeginProperty Column32 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column33 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column34 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column35 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column36 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column37 
            ColumnWidth     =   1530.142
         EndProperty
         BeginProperty Column38 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column39 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column40 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column41 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column42 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column43 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column44 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column45 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column46 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column47 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column48 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column49 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column50 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column51 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column52 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column53 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column54 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column55 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column56 
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column57 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column58 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column59 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column60 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column61 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column62 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column63 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column64 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column65 
            ColumnWidth     =   1365.165
         EndProperty
         BeginProperty Column66 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column67 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column68 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column69 
            ColumnWidth     =   1425.26
         EndProperty
         BeginProperty Column70 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column71 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column72 
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column73 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column74 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column75 
            ColumnWidth     =   1124.787
         EndProperty
         BeginProperty Column76 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column77 
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column78 
            ColumnWidth     =   900.284
         EndProperty
      EndProperty
   End
   Begin Threed.SSPanel pbxBottom 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   9240
      Width           =   15015
      _Version        =   65536
      _ExtentX        =   26485
      _ExtentY        =   1482
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
      BevelOuter      =   0
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         DisabledPicture =   "frmSearchRcp.frx":57FB
         DownPicture     =   "frmSearchRcp.frx":643D
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
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchRcp.frx":707F
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         DisabledPicture =   "frmSearchRcp.frx":7CC1
         DownPicture     =   "frmSearchRcp.frx":8903
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
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchRcp.frx":9545
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         DisabledPicture =   "frmSearchRcp.frx":A187
         DownPicture     =   "frmSearchRcp.frx":ADC9
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
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmSearchRcp.frx":BA0B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.CommandButton cmdCreateNew 
         Caption         =   " Create New"
         DisabledPicture =   "frmSearchRcp.frx":C64D
         DownPicture     =   "frmSearchRcp.frx":C98F
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
         Picture         =   "frmSearchRcp.frx":CCD1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.Timer tmrScreen 
         Interval        =   250
         Left            =   13800
         Top             =   480
      End
      Begin MSAdodcLib.Adodc adoRecipes 
         Height          =   375
         Left            =   11520
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
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
         RecordSource    =   "SELECT * FROM [MasterRecipe] ORDER BY [Number] ASC"
         Caption         =   "Recipes"
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
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "message"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   4665
         TabIndex        =   2
         Top             =   120
         Width           =   6495
      End
   End
End
Attribute VB_Name = "frmSearchRcp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 158'''''''''' Form SearchRcp.frm '''''''''''''''''''
Option Explicit
'
Private RecipeMode As Integer            ' 0=master; 1=station
Private sPath As String
Private rsCrit As String
Private RcpSelDest As Integer
Private antiRepeatDelete As Boolean
Private searchRcpMsg As String
Private searchRcpMsgColor As Long

Public Sub ChgRecipeMode(ByVal NewMode As Integer)
    ' 0=master; 1=station
    RecipeMode = IIf((NewMode = 0 Or NewMode = 1), NewMode, 0)
    Select Case RecipeMode
        Case MASTERMODE
            ' clear button
            cmdClear.Visible = True
            ' create new button
            cmdCreateNew.Visible = True
            ' delete button
            cmdDelete.Visible = True
        Case STATIONMODE
            ' clear button
            cmdClear.Visible = False
            ' create new button
            cmdCreateNew.Visible = False
            ' delete button
            cmdDelete.Visible = False
    End Select
End Sub

Private Sub adoRecipes_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    antiRepeatDelete = False
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
End Sub

Public Sub ChgSelectionDestination(ByVal NewDest As Integer)
    ' 1=course; 2=recipe
    RcpSelDest = IIf((NewDest = rcpdestCourse Or NewDest = rcpdestRecipe), NewDest, rcpdestRecipe)
End Sub

Private Sub cmdCreateNew_Click()
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
    NewRcp
End Sub

Private Sub cmdClear_Click()
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
    ClearRcp
End Sub

Private Sub cmdDelete_Click()
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
    DeleteRcp
End Sub

Private Sub Xit()
    Unload frmSearchRcp
    Set frmSearchRcp = Nothing
End Sub

Private Sub cmdSelect_Click()
Dim recnum As Integer
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
    If Not adoRecipes.Recordset.BOF Then
        recnum = CInt(dgRecipes.Columns(0).CellValue(dgRecipes.GetBookmark(0)))
        Select Case RcpSelDest
            Case rcpdestCourse
                frmCourses.Show
                frmCourses.LoadRcpNum CInt(recnum)
            Case rcpdestRecipe
                frmRecipe.Show
                frmRecipe.LoadNewRcp CInt(recnum)
        End Select
        Unload frmSearchRcp
        Set frmSearchRcp = Nothing
    End If
End Sub

Private Sub dgRecipes_HeadClick(ByVal ColIndex As Integer)
    DisplayData (ColIndex + 1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
      Unload Me
      Set frmSearchRcp = Nothing
    End If
End Sub

Private Sub Form_Load()
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 158, 2
Dim flag1 As Boolean
Dim flag2 As Boolean

    KeyPreview = True
    
    flag1 = CheckPass("P", False) And CheckPass("7", False)
    flag2 = CheckPass("P", False) And (CheckPass("8", False) Or CheckPass("7", False))
    cmdSelect.Visible = IIf(flag2, True, False)
    cmdClear.Visible = IIf(flag1, True, False)
    cmdCreateNew.Visible = IIf(flag2, True, False)
    cmdDelete.Visible = IIf(flag2, True, False)
    
    dgRecipes.AllowRowSizing = False
    
    DisplayData 1
    
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub DisplayData(sortCol As Integer)
    ' Select & Sort
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = ""
    Select Case sortCol
        Case 1
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Number] ASC"
        Case 2
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Name] ASC"
        Case 3
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Cycles] DESC"
        Case 4
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Load_Method] DESC"
        Case 5
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [UseHiRangeMFC] DESC"
        Case 6
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [NitrogenFlow] DESC"
        Case 7
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Load_Rate] DESC"
        Case 8
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Mix_Percent] DESC"
        Case 9
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [WC_Mult] DESC"
        Case 10
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [EPAFill] DESC"
        Case 11
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Load_Wt] DESC"
        Case 12
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LoadBreakthrough] DESC"
        Case 13
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [FidBreakthrough] DESC"
        Case 14
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [FIDmg] DESC"
        Case 15
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Load_Time] DESC"
        Case 16
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Purge_Flow] DESC"
        Case 17
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Purge_Can_Vol] DESC"
        Case 18
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Purge_Time] DESC"
        Case 19
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [UseAuxScale] DESC"
        Case 20
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PurgeAuxCan] DESC"
        Case 21
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [AuxScaleNo] DESC"
        Case 22
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PauseAfterLeak] DESC"
        Case 23
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PauseLeakTime] DESC"
        Case 24
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PauseAfterLoad] DESC"
        Case 25
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PauseLoadTime] DESC"
        Case 26
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PauseAfterPurge] DESC"
        Case 27
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PausePurgeTime] DESC"
        Case 28
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PrimaryScale] DESC"
        Case 29
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PrimaryScaleNo] DESC"
        Case 30
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [TargetConcentration] DESC"
        Case 31
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [DwellTime] DESC"
        Case 32
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LeakCheck] DESC"
        Case 33
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LeakPrimary] DESC"
        Case 34
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LeakAux] DESC"
        Case 35
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [UseAnalyser] DESC"
        Case 36
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Run15BWCycle] DESC"
        Case 37
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [MaxLoadTime] DESC"
        Case 38
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [IDLoad] DESC"
        Case 39
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LoadL] DESC"
        Case 40
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [IDPurge] DESC"
        Case 41
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PurgeL] DESC"
        Case 42
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [IDVent] DESC"
        Case 43
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [VentL] DESC"
        Case 44
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LoadV] DESC"
        Case 45
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [PurgeV] DESC"
        Case 46
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [VentV] DESC"
        Case 47
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LiveFuel] DESC"
        Case 48
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LiveFuelChgAuto] DESC"
        Case 49
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [LiveFuelChgFreq] DESC"
        Case 50
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [ADF_Heater] DESC"
        Case 51
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [ADF_HeaterSP] DESC"
        Case 52
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [StartMethod] DESC"
        Case 53
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [StartDelay] DESC"
        Case 54
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [StartDate] DESC"
        Case Else
            sortCol = 1
            rsCrit = "SELECT * FROM [MasterRecipe] ORDER BY [Number] ASC"
    End Select
    adoRecipes.RecordSource = rsCrit
    adoRecipes.Refresh

    If adoRecipes.Recordset.BOF Then
        dgRecipes.Caption = " No Defined Recipess"
        ' Set column properties
        dgRecipes.Columns(0).Width = 760
        dgRecipes.Columns(1).Width = 4000
        dgRecipes.Columns(2).Width = 760
        dgRecipes.Columns(3).Width = 1250
        dgRecipes.Columns(10).Width = 1650
        dgRecipes.Columns(11).Width = 1500
        dgRecipes.Columns(49).Width = 2400
        cmdClear.Enabled = False
        cmdDelete.Enabled = False
        cmdSelect.Enabled = False
    Else
        ' Display number of recipes found
        adoRecipes.Recordset.GetRows
        Select Case adoRecipes.Recordset.RecordCount
            Case 0
                dgRecipes.Caption = " No Defined Recipes"
                cmdClear.Enabled = False
                cmdSelect.Enabled = False
                cmdSelect.Enabled = False
            Case 1
                dgRecipes.Caption = Format(adoRecipes.Recordset.RecordCount, "###0") & " Defined Recipe"
                cmdClear.Enabled = False
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
            Case Else
                dgRecipes.Caption = Format(adoRecipes.Recordset.RecordCount, "###0") & " Defined Recipes"
                cmdClear.Enabled = True
                cmdDelete.Enabled = True
                cmdSelect.Enabled = True
        End Select
        dgRecipes.Refresh
        ' Set column properties
        dgRecipes.Columns(0).Width = 760
        dgRecipes.Columns(1).Width = 4000
        dgRecipes.Columns(2).Width = 760
        dgRecipes.Columns(3).Width = 1250
        dgRecipes.Columns(10).Width = 1650
        dgRecipes.Columns(11).Width = 1500
        dgRecipes.Columns(49).Width = 2400
        
        ' move pointer to first row
        adoRecipes.Recordset.MoveFirst
        
        ' make the Left-Most column the Sorted-By column
        dgRecipes.LeftCol = IIf(sortCol > 51, 51, sortCol - 1)
    End If
End Sub

Private Sub DeleteRcp()
SetErrModule 158, 31
If UseLocalErrorHandler Then On Error GoTo localhandler
    If Not antiRepeatDelete Then
        If adoRecipes.Recordset.BOF Then
            searchRcpMsgColor = MEDRED
            searchRcpMsg = "No Recipe Data Available"
        Else
        
            If IsNull(dgRecipes.Columns(0).CellValue(dgRecipes.GetBookmark(0))) Or IsEmpty(dgRecipes.Columns(0).CellValue(dgRecipes.GetBookmark(0))) Then
                ' Report an error
                searchRcpMsgColor = MEDRED
                searchRcpMsg = "Invalid Recipe Number"
                Exit Sub
            End If
            
            adoRecipes.Recordset.Delete
            searchRcpMsgColor = Message_ForeColor
            searchRcpMsg = "Recipe Deleted"
            antiRepeatDelete = True
           
        End If
    End If
Exit Sub
localhandler:
    searchRcpMsgColor = MEDRED
    searchRcpMsg = "Unable to Delete Recipe"
End Sub

Private Sub ClearRcp()
Dim iAux As Integer

    SetErrModule 158, 3
    If UseLocalErrorHandler Then On Error GoTo localhandler
    ' Clearing Recipe
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = "Clearing Recipe.. Please Wait"
    adoRecipes.RecordSource = "SELECT * FROM [MasterRecipe] with [Number] = " & (dgRecipes.Columns(0).CellValue(dgRecipes.GetBookmark(0))) & "  ORDER BY [Number] ASC"
'    adoRecipes.Recordset.MoveLast
    
    ' Blank Recipe
    adoRecipes.Recordset.Fields("Name").Value = "undefined"
'    adoRecipes.Recordset.Fields("Number").Value = CInt(0)
    
    adoRecipes.Recordset.Fields("CycleType").Value = CyclePurgeLoad
    
    adoRecipes.Recordset.Fields("Load_Method").Value = NOLOAD
    adoRecipes.Recordset.Fields("Load_MethodDesc").Value = "No Load"
    adoRecipes.Recordset.Fields("UseHiRangeMFC").Value = False
    adoRecipes.Recordset.Fields("UseLoadRatePID").Value = False
    adoRecipes.Recordset.Fields("NitrogenFlow").Value = CSng(0)
'    adoRecipes.Recordset.Fields("NitrogenFlowSave").Value = CSng(0)
    adoRecipes.Recordset.Fields("Load_Rate").Value = 0
'    adoRecipes.Recordset.Fields("Load_RateSave").Value = 0
    adoRecipes.Recordset.Fields("Mix_Percent").Value = 0
    adoRecipes.Recordset.Fields("WC_Mult").Value = 0
    adoRecipes.Recordset.Fields("EPAFill").Value = 0
    adoRecipes.Recordset.Fields("Load_Wt").Value = 0
    adoRecipes.Recordset.Fields("LoadBreakthrough").Value = 0
    adoRecipes.Recordset.Fields("FIDmg").Value = CSng(0)
    adoRecipes.Recordset.Fields("Load_Time").Value = 0
    
    adoRecipes.Recordset.Fields("Purge_Method").Value = NOPURGE
    adoRecipes.Recordset.Fields("Purge_MethodDesc").Value = "No Purge"
    adoRecipes.Recordset.Fields("Purge_Flow").Value = 0
    adoRecipes.Recordset.Fields("Purge_Time").Value = 0
    adoRecipes.Recordset.Fields("Purge_AuxTime").Value = 0
    adoRecipes.Recordset.Fields("Purge_Can_Vol").Value = 0
    adoRecipes.Recordset.Fields("Purge_ProfileNumber").Value = 0
    adoRecipes.Recordset.Fields("Purge_TargetMode").Value = 0
    adoRecipes.Recordset.Fields("Purge_TargetModeDesc").Value = "Continuous Purge"
    adoRecipes.Recordset.Fields("Purge_TargetWC").Value = 0
    adoRecipes.Recordset.Fields("Purge_TargetWeight").Value = 0
    adoRecipes.Recordset.Fields("Purge_MaxVolumes").Value = 0
    adoRecipes.Recordset.Fields("Purge_TargetPurge").Value = 0
    adoRecipes.Recordset.Fields("Purge_TargetPause").Value = 0
    
    adoRecipes.Recordset.Fields("UseAuxScale").Value = False
    adoRecipes.Recordset.Fields("PurgeAuxCan").Value = False
    adoRecipes.Recordset.Fields("AuxScaleNo").Value = 0
    adoRecipes.Recordset.Fields("PauseLeakTime").Value = 0
    adoRecipes.Recordset.Fields("PauseLoadTime").Value = 0
    adoRecipes.Recordset.Fields("PausePurgeTime").Value = 0
    adoRecipes.Recordset.Fields("UsePriScale").Value = False
    adoRecipes.Recordset.Fields("PriScaleNo").Value = 0
    adoRecipes.Recordset.Fields("PauseAfterLeak").Value = False
    adoRecipes.Recordset.Fields("PauseAfterLoad").Value = False
    adoRecipes.Recordset.Fields("PauseAfterPurge").Value = False
    adoRecipes.Recordset.Fields("TargetConcentration").Value = False
    adoRecipes.Recordset.Fields("DwellTime").Value = 0
    adoRecipes.Recordset.Fields("LeakCheck").Value = False
    adoRecipes.Recordset.Fields("LeakPrimary").Value = False
    adoRecipes.Recordset.Fields("LeakAux").Value = False
    adoRecipes.Recordset.Fields("UseAnalyzer").Value = False
    adoRecipes.Recordset.Fields("MaxLoadTime").Value = 0
    
    adoRecipes.Recordset.Fields("IDLoad").Value = 0
    adoRecipes.Recordset.Fields("LoadL").Value = 0
    adoRecipes.Recordset.Fields("LoadV").Value = 0
    adoRecipes.Recordset.Fields("IDPurge").Value = 0
    adoRecipes.Recordset.Fields("PurgeL").Value = 0
    adoRecipes.Recordset.Fields("PurgeV").Value = 0
    adoRecipes.Recordset.Fields("IDVent").Value = 0
    adoRecipes.Recordset.Fields("VentL").Value = 0
    adoRecipes.Recordset.Fields("VentV").Value = 0
    
    adoRecipes.Recordset.Fields("LiveFuel").Value = False
    adoRecipes.Recordset.Fields("LiveFuelChgAuto").Value = False
    adoRecipes.Recordset.Fields("LiveFuelChgFreq").Value = 0
    adoRecipes.Recordset.Fields("ADF_Heater").Value = False
    adoRecipes.Recordset.Fields("ADF_HeaterSP").Value = 0
    
    adoRecipes.Recordset.Fields("StartMethod").Value = STARTNOW
    adoRecipes.Recordset.Fields("StartMethodDesc").Value = "Start Without Delay"
    adoRecipes.Recordset.Fields("StartDelay").Value = 0
    adoRecipes.Recordset.Fields("StartDate").Value = Now()
                
    ' end method
    adoRecipes.Recordset.Fields("EndMethod").Value = 0
    adoRecipes.Recordset.Fields("EndMethodDesc").Value = "End after x cycles"
    adoRecipes.Recordset.Fields("Cycles").Value = CInt(0)
    adoRecipes.Recordset.Fields("EndWeightTolerance").Value = CInt(0)
    adoRecipes.Recordset.Fields("EndConsecutiveCycles").Value = CInt(0)
    adoRecipes.Recordset.Fields("EndMaximumCycles").Value = CInt(0)
    adoRecipes.Recordset.Fields("EndMinimumCycles").Value = CInt(0)

    adoRecipes.Recordset.Fields("AuxOutputs").Value = False
    adoRecipes.Recordset.Fields("AuxOutput1_Load").Value = False
    adoRecipes.Recordset.Fields("AuxOutput1_Purge").Value = False
    adoRecipes.Recordset.Fields("AuxOutput2_Load").Value = False
    adoRecipes.Recordset.Fields("AuxOutput2_Purge").Value = False
    adoRecipes.Recordset.Fields("AuxOutput3_Load").Value = False
    adoRecipes.Recordset.Fields("AuxOutput3_Purge").Value = False
    adoRecipes.Recordset.Fields("AuxOutput4_Load").Value = False
    adoRecipes.Recordset.Fields("AuxOutput4_Purge").Value = False
   
    adoRecipes.Recordset.Update
    adoRecipes.RecordSource = "SELECT * FROM [MasterRecipe] ORDER BY [Number] ASC"
    dgRecipes.Refresh
                    
    searchRcpMsgColor = Message_ForeColor
    searchRcpMsg = "Recipe Cleared"
    
ResetErrModule
Exit Sub

localhandler:
    searchRcpMsgColor = MEDRED
    searchRcpMsg = "Unable to Clear Recipe"
End Sub

Private Sub NewRcp()
Dim iRcp As Integer
Dim rcpnum As Integer

    rcpnum = 0
    For iRcp = 1 To MAX_RCP
        If rcpnum = 0 Then
            If Not IsDefined(iRcp, adoRecipes.Recordset) Then
                rcpnum = iRcp
            End If
        End If
    Next iRcp
    
    If rcpnum > 0 Then
        frmRecipe.Show
        frmRecipe.LoadNewRcp CInt(rcpnum)
        Unload frmSearchRcp
        Set frmSearchRcp = Nothing
    Else
        searchRcpMsgColor = MEDRED
        searchRcpMsg = "No undefined Master Recipe"
    End If
End Sub

Private Function IsDefined(ByVal iNum As Integer, ByRef rS As ADODB.Recordset) As Boolean
Dim flag As Boolean

    flag = False
    
    With rS
    
        If Not .BOF Or Not .EOF Then
        
            .MoveLast
            .MoveFirst
            .MoveLast
            
            Do Until (.BOF Or flag)
                If ((iNum = .Fields("Number").Value) And (.Fields("Name").Value <> "undefined") And (Len(Trim(.Fields("Name").Value)) > 0)) Then
                    flag = True
                End If
                .MovePrevious
            Loop
        
        End If
    
    End With
    
    IsDefined = flag
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    Xit
End Sub

Private Sub tmrScreen_Timer()
    lblMessage.ForeColor = searchRcpMsgColor
    lblMessage.Caption = searchRcpMsg
End Sub
