VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuration Options Screen"
   ClientHeight    =   9060
   ClientLeft      =   135
   ClientTop       =   495
   ClientWidth     =   9330
   ClipControls    =   0   'False
   Icon            =   "frmConfi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9060
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pbCapture 
      Height          =   2295
      Left            =   10800
      ScaleHeight     =   2235
      ScaleWidth      =   2955
      TabIndex        =   157
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "frmConfi.frx":57E2
      DownPicture     =   "frmConfi.frx":6424
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
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmConfi.frx":7066
      Style           =   1  'Graphical
      TabIndex        =   123
      ToolTipText     =   "Save Configuration Values"
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      DisabledPicture =   "frmConfi.frx":7CA8
      DownPicture     =   "frmConfi.frx":88EA
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
      Left            =   1440
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmConfi.frx":952C
      Style           =   1  'Graphical
      TabIndex        =   122
      ToolTipText     =   "Print Configuration Values"
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Quit"
      DisabledPicture =   "frmConfi.frx":A16E
      DownPicture     =   "frmConfi.frx":ADB0
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
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmConfi.frx":B9F2
      Style           =   1  'Graphical
      TabIndex        =   121
      ToolTipText     =   "Quit"
      Top             =   8040
      UseMaskColor    =   -1  'True
      Width           =   840
   End
   Begin TabDlg.SSTab cfgtabs 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   9
      Tab             =   7
      TabsPerRow      =   9
      TabHeight       =   609
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Leak Check"
      TabPicture(0)   =   "frmConfi.frx":C634
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SelectLkChkFailResp"
      Tab(0).Control(1)=   "txtLCIntvl"
      Tab(0).Control(2)=   "txtLCTime"
      Tab(0).Control(3)=   "txtLCMinDelay"
      Tab(0).Control(4)=   "txtPressureDecay"
      Tab(0).Control(5)=   "txtLCSetPoint"
      Tab(0).Control(6)=   "lblLeakErrResponse"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label9(1)"
      Tab(0).Control(9)=   "Label36"
      Tab(0).Control(10)=   "Label32"
      Tab(0).Control(11)=   "lblLCSetPoint"
      Tab(0).Control(12)=   "Label20"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Load"
      TabPicture(1)   =   "frmConfi.frx":C650
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "WaterBathTemperatureControl"
      Tab(1).Control(1)=   "txtWaterBathTol"
      Tab(1).Control(2)=   "txtLoadTimeLimit"
      Tab(1).Control(3)=   "txtButaneMassLimit"
      Tab(1).Control(4)=   "txtLoadSettleTime"
      Tab(1).Control(5)=   "txtORVRBtnTol"
      Tab(1).Control(6)=   "txtORVRNitTol"
      Tab(1).Control(7)=   "txtFuelTempTol"
      Tab(1).Control(8)=   "txtLfvTol"
      Tab(1).Control(9)=   "txtNitTol"
      Tab(1).Control(10)=   "txtLoadTotIntvl"
      Tab(1).Control(11)=   "txtLoLimLoad"
      Tab(1).Control(12)=   "txtCanventOvr"
      Tab(1).Control(13)=   "txtLoadPressure"
      Tab(1).Control(14)=   "txtNitrogenPurgeTime"
      Tab(1).Control(15)=   "txtLoadIntvl"
      Tab(1).Control(16)=   "txtLoadTotal"
      Tab(1).Control(17)=   "txtMixRatio"
      Tab(1).Control(18)=   "txtBtnTol"
      Tab(1).Control(19)=   "lblWaterBathUnits"
      Tab(1).Control(20)=   "lblWaterBathDesc"
      Tab(1).Control(21)=   "lblLoadTimeLimit"
      Tab(1).Control(22)=   "lblLoadTimeLimitUnits"
      Tab(1).Control(23)=   "lblButaneMassLimit"
      Tab(1).Control(24)=   "lblButaneMassLimitUnits"
      Tab(1).Control(25)=   "lblLoadSettleTime"
      Tab(1).Control(26)=   "lblORVRButFlowUnits"
      Tab(1).Control(27)=   "lblORVRNitFlowUnits"
      Tab(1).Control(28)=   "lblORVRButFlowTol"
      Tab(1).Control(29)=   "lblORVRNitFlowTol"
      Tab(1).Control(30)=   "lblFuelTempTol"
      Tab(1).Control(31)=   "lblFuelTempUnits"
      Tab(1).Control(32)=   "lblFuelFlowUnits"
      Tab(1).Control(33)=   "lblFuelFlowTol"
      Tab(1).Control(34)=   "lblNitFlowUnits"
      Tab(1).Control(35)=   "lblNitFlowTol"
      Tab(1).Control(36)=   "Label48"
      Tab(1).Control(37)=   "Label47"
      Tab(1).Control(38)=   "Label45"
      Tab(1).Control(39)=   "Label44"
      Tab(1).Control(40)=   "lblCanventUnits"
      Tab(1).Control(41)=   "lblCanventDescr"
      Tab(1).Control(42)=   "lblButFlowUnits"
      Tab(1).Control(43)=   "lblLoadPressureUnits"
      Tab(1).Control(44)=   "lblLoadPressure"
      Tab(1).Control(45)=   "lblNitrogenPurgeTimeUnits"
      Tab(1).Control(46)=   "lblNitrogenPurgeTime"
      Tab(1).Control(47)=   "Label53"
      Tab(1).Control(48)=   "lblButFlowTol"
      Tab(1).Control(49)=   "Label54"
      Tab(1).Control(50)=   "lblMixTol"
      Tab(1).Control(51)=   "Label55"
      Tab(1).Control(52)=   "lblMixUnits"
      Tab(1).Control(53)=   "Label56"
      Tab(1).ControlCount=   54
      TabCaption(2)   =   "Purge"
      TabPicture(2)   =   "frmConfi.frx":C66C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtOvenTempTol"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkDryAirPurge"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkPosPressPurge"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtOvenBand"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "txtPurgeDpHiLimit"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtPurgeSettleTime"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtPurgeTotIntvl"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtLoLimPurge"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtMoistureTarget"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtTempTarget"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtPurgeTotal"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtMoistureTol"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtTempTol"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtPurgeIntvl"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "txtPurgeTol"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "lblOvenTempTol"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "lblOvenTempUnits"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "lblOvenBand"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "lblOvenBandUnits"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "lblPurgeDpHiLimit"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "lblPurgeDpHiLimitUnits"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "lblPurgeSettleTime"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Label46"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Label13"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Label26"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Label15"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "lblMoistTargetUnits"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "lblMoistureTarget"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "lblTempTargetUnits"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "lblTempTarget"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Label22"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Label21"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "lblMoistTolUnits"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "lblTempTolUnits"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Label24"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Label25"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "lblMoistureTol"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "lblTempTol"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Label31"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Label33"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).ControlCount=   40
      TabCaption(3)   =   "Job"
      TabPicture(3)   =   "frmConfi.frx":C688
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtFileName"
      Tab(3).Control(1)=   "optDbfBackup"
      Tab(3).Control(2)=   "txtDbfBackupPath"
      Tab(3).Control(3)=   "cboReportName2"
      Tab(3).Control(4)=   "txtRptBackupPath"
      Tab(3).Control(5)=   "optRptBackup"
      Tab(3).Control(6)=   "cboReportName3"
      Tab(3).Control(7)=   "cboReportName1"
      Tab(3).Control(8)=   "txtHeading"
      Tab(3).Control(9)=   "txtHeading2"
      Tab(3).Control(10)=   "Label34"
      Tab(3).Control(11)=   "Label35"
      Tab(3).Control(12)=   "Label37"
      Tab(3).Control(13)=   "Label38"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Process"
      TabPicture(4)   =   "frmConfi.frx":C6A4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtDefaultIntvl"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "txtOnDutyMult(1)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "txtOffDutyMult(1)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "txtPgain(2)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txtOOTtime"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txtTimeoutDuration(2)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "txtTimeoutDuration(1)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "txtIgain(2)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "txtInTolDuration(2)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "txtInTolDuration(1)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "txtDoorOpenDelay"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "txtUPSOpenDelay"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Label2"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Label1"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Label40"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "lblTimeoutDuration(2)"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "lblTimeoutDuration(1)"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "lblIgain(2)"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "lblPgain(2)"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "lblInTolDuration(2)"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "lblPidControl(2)"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "lblInTolDuration(1)"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "lblDoorOpenDelay"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "lblUPSOpenDelay"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "lblPidControl(1)"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "lblOffDutyMult(1)"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "lblOnDutyMult(1)"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).ControlCount=   27
      TabCaption(5)   =   "System"
      TabPicture(5)   =   "frmConfi.frx":C6C0
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "txtRemStatusLogInterval"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "optLogTempRhVerbose"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "txtLogTempRhInterval"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "SelectUserName"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "txtJobRecs"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "txtEventRecs"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "lblRemStatusLogIntervalUnits"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lblRemStatusLogInterval"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "lblLogTempRhIntervalUnits"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "lblLogTempRhInterval"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "lblAutoLogon"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "lblJobRecs"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "lblEventRecs"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).ControlCount=   13
      TabCaption(6)   =   "Reporting"
      TabPicture(6)   =   "frmConfi.frx":C6DC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "frmGenOptions"
      Tab(6).Control(1)=   "frmEotOptions"
      Tab(6).ControlCount=   2
      TabCaption(7)   =   "OOT Response"
      TabPicture(7)   =   "frmConfi.frx":C6F8
      Tab(7).ControlEnabled=   -1  'True
      Tab(7).Control(0)=   "frmOotResp(9)"
      Tab(7).Control(0).Enabled=   0   'False
      Tab(7).Control(1)=   "frmOotResp(8)"
      Tab(7).Control(1).Enabled=   0   'False
      Tab(7).Control(2)=   "frmOotResp(4)"
      Tab(7).Control(2).Enabled=   0   'False
      Tab(7).Control(3)=   "frmOotResp(5)"
      Tab(7).Control(3).Enabled=   0   'False
      Tab(7).Control(4)=   "frmOotResp(3)"
      Tab(7).Control(4).Enabled=   0   'False
      Tab(7).Control(5)=   "frmOotResp(2)"
      Tab(7).Control(5).Enabled=   0   'False
      Tab(7).Control(6)=   "frmOotResp(1)"
      Tab(7).Control(6).Enabled=   0   'False
      Tab(7).Control(7)=   "frmOotResp(7)"
      Tab(7).Control(7).Enabled=   0   'False
      Tab(7).Control(8)=   "frmOotResp(6)"
      Tab(7).Control(8).Enabled=   0   'False
      Tab(7).Control(9)=   "frmOotResp(10)"
      Tab(7).Control(9).Enabled=   0   'False
      Tab(7).Control(10)=   "frmOotResp(11)"
      Tab(7).Control(10).Enabled=   0   'False
      Tab(7).Control(11)=   "frmOotResp(12)"
      Tab(7).Control(11).Enabled=   0   'False
      Tab(7).Control(12)=   "frmOotResp(13)"
      Tab(7).Control(12).Enabled=   0   'False
      Tab(7).ControlCount=   13
      TabCaption(8)   =   "AutoDrainFill"
      TabPicture(8)   =   "frmConfi.frx":C714
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "txtFuelStorageLeakRate"
      Tab(8).Control(1)=   "txtVaporGenLeakRate"
      Tab(8).Control(2)=   "txtFuelStorageFillShutoff"
      Tab(8).Control(3)=   "txtFuelStorageFillTimeout"
      Tab(8).Control(4)=   "txtFuelStorageFillDelay"
      Tab(8).Control(5)=   "txtFuelStorageDrainShutoff"
      Tab(8).Control(6)=   "txtFuelStorageDrainTimeout"
      Tab(8).Control(7)=   "txtFuelStorageDrainDelay"
      Tab(8).Control(8)=   "txtFuelStorageTankVol"
      Tab(8).Control(9)=   "txtVaporGenTankVol"
      Tab(8).Control(10)=   "txtLoadRate_Pgain"
      Tab(8).Control(11)=   "txtLoadRate_Igain"
      Tab(8).Control(12)=   "txtLiveFuelChgPurgeFillDelay"
      Tab(8).Control(13)=   "txtLiveFuelChgPurgeTimeout"
      Tab(8).Control(14)=   "txtLiveFuelChgPurgeDrainDelay"
      Tab(8).Control(15)=   "txtLiveFuelChgDrainDelay"
      Tab(8).Control(16)=   "txtLiveFuelChgFillDelay"
      Tab(8).Control(17)=   "txtLiveFuelChgDrainTimeout"
      Tab(8).Control(18)=   "txtLiveFuelChgFillTimeout"
      Tab(8).Control(19)=   "txtLiveFuelChgHeaterTimeout"
      Tab(8).Control(20)=   "txtLiveFuelChgFillShutoff"
      Tab(8).Control(21)=   "txtLiveFuelChgDrainShutoff"
      Tab(8).Control(22)=   "frmStnSelection"
      Tab(8).Control(23)=   "lblFuelStorageLeakRate"
      Tab(8).Control(24)=   "lblFuelStorageLeakRateUnits"
      Tab(8).Control(25)=   "lblVaporGenLeakRate"
      Tab(8).Control(26)=   "lblVaporGenLeakRateUnits"
      Tab(8).Control(27)=   "lblLoadByPID"
      Tab(8).Control(28)=   "lblFST_Shutoff2"
      Tab(8).Control(29)=   "lblFST_Shutoff"
      Tab(8).Control(30)=   "lblFST_Timeout"
      Tab(8).Control(31)=   "lblFST_Timeout2"
      Tab(8).Control(32)=   "lblFST_Delay"
      Tab(8).Control(33)=   "lblFST_Delay2"
      Tab(8).Control(34)=   "lblStorageTank"
      Tab(8).Control(35)=   "lblFST_Fill"
      Tab(8).Control(36)=   "lblFST_Drain"
      Tab(8).Control(37)=   "lblVaportank"
      Tab(8).Control(38)=   "lblFuelStorageTankVol2"
      Tab(8).Control(39)=   "lblFuelStorageTankVol"
      Tab(8).Control(40)=   "lblVaporGenTankVol2"
      Tab(8).Control(41)=   "lblVaporGenTankVol"
      Tab(8).Control(42)=   "lblLoadRate_Pgain"
      Tab(8).Control(43)=   "lblLoadRate_Igain"
      Tab(8).Control(44)=   "lblADF_PurgeFillDelay2"
      Tab(8).Control(45)=   "lblADF_PurgeFillDelay"
      Tab(8).Control(46)=   "lblADF_PurgeTimeout"
      Tab(8).Control(47)=   "lblADF_PurgeTimeout2"
      Tab(8).Control(48)=   "lblADF_PurgeDrainDelay"
      Tab(8).Control(49)=   "lblADF_PurgeDrainDelay2"
      Tab(8).Control(50)=   "lblADF_Drain"
      Tab(8).Control(51)=   "lblADF_Fill"
      Tab(8).Control(52)=   "lblADF_Delay2"
      Tab(8).Control(53)=   "lblADF_Delay"
      Tab(8).Control(54)=   "lblADF_Timeout2"
      Tab(8).Control(55)=   "lblADF_Timeout"
      Tab(8).Control(56)=   "lblADF_HeaterTimeout"
      Tab(8).Control(57)=   "lblADF_HeaterTimeout2"
      Tab(8).Control(58)=   "lblADF_Shutoff"
      Tab(8).Control(59)=   "lblADF_Shutoff2"
      Tab(8).ControlCount=   60
      Begin VB.TextBox txtRemStatusLogInterval 
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
         Left            =   -70785
         MaxLength       =   6
         TabIndex        =   288
         Text            =   "0"
         ToolTipText     =   "30 to 60 seconds"
         Top             =   4200
         Width           =   735
      End
      Begin VB.ComboBox WaterBathTemperatureControl 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "frmConfi.frx":C730
         Left            =   -74760
         List            =   "frmConfi.frx":C73D
         Style           =   2  'Dropdown List
         TabIndex        =   287
         ToolTipText     =   "Select Response to Purge Flow Rate OutOfTolerance condition"
         Top             =   4200
         Width           =   4215
      End
      Begin VB.TextBox txtOvenTempTol 
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
         Left            =   -66555
         MaxLength       =   3
         TabIndex        =   284
         Text            =   "9.0"
         ToolTipText     =   "0 to 100 Degrees C"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox txtWaterBathTol 
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
         Left            =   -71205
         MaxLength       =   4
         TabIndex        =   281
         Text            =   "0"
         ToolTipText     =   "2 to 15 deg F"
         Top             =   2280
         Width           =   735
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "WaterBath OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Index           =   13
         Left            =   0
         TabIndex        =   278
         Top             =   3840
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   13
            ItemData        =   "frmConfi.frx":C795
            Left            =   1230
            List            =   "frmConfi.frx":C7A2
            Style           =   2  'Dropdown List
            TabIndex        =   279
            ToolTipText     =   "Select Response to WaterBath Temperature OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   13
            Left            =   240
            TabIndex        =   280
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Purge Oven OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   810
         Index           =   12
         Left            =   6150
         TabIndex        =   275
         Top             =   3840
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   12
            ItemData        =   "frmConfi.frx":C7BD
            Left            =   1230
            List            =   "frmConfi.frx":C7CA
            Style           =   2  'Dropdown List
            TabIndex        =   276
            ToolTipText     =   "Select Response to Purge Oven Temperature OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   240
            TabIndex        =   277
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.CheckBox chkDryAirPurge 
         Caption         =   " Use Dry Air Purge?"
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
         Left            =   -74880
         TabIndex        =   273
         Top             =   2460
         Width           =   3000
      End
      Begin VB.CheckBox chkPosPressPurge 
         Caption         =   " Use Positive Pressure Purge?"
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
         Left            =   -74880
         TabIndex        =   272
         Top             =   2220
         Width           =   3000
      End
      Begin VB.TextBox txtOvenBand 
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
         Left            =   -66555
         MaxLength       =   6
         TabIndex        =   269
         Text            =   "9.0"
         ToolTipText     =   "0 to 100 Degrees C"
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtDefaultIntvl 
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
         Left            =   -70890
         MaxLength       =   6
         TabIndex        =   266
         Text            =   "0"
         ToolTipText     =   "1 to 900 seconds (Does not apply to Leakcheck or Purge or Load)"
         Top             =   1080
         Width           =   735
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Storage Tank Level OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   810
         Index           =   11
         Left            =   3075
         TabIndex        =   263
         Top             =   1410
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   11
            ItemData        =   "frmConfi.frx":C7E5
            Left            =   1230
            List            =   "frmConfi.frx":C7F2
            Style           =   2  'Dropdown List
            TabIndex        =   264
            ToolTipText     =   "Select Response to Fuel Tank Level OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   240
            TabIndex        =   265
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.TextBox txtFuelStorageLeakRate 
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
         Left            =   -71445
         MaxLength       =   5
         TabIndex        =   260
         Text            =   "0.00"
         ToolTipText     =   "Fuel Storage Tank Level Tolerance in gallons  (0.01-10)"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtVaporGenLeakRate 
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
         Left            =   -66720
         MaxLength       =   5
         TabIndex        =   257
         Text            =   "0.00"
         ToolTipText     =   "Vapor Generator Tank Level Tolerance in gallons  (0.01-10)"
         Top             =   3720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Fuel Tank Level OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   810
         Index           =   10
         Left            =   3075
         TabIndex        =   254
         Top             =   2220
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            ItemData        =   "frmConfi.frx":C80D
            Left            =   1230
            List            =   "frmConfi.frx":C81A
            Style           =   2  'Dropdown List
            TabIndex        =   255
            ToolTipText     =   "Select Response to Fuel Tank Level OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   240
            TabIndex        =   256
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Air Temp OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   810
         Index           =   6
         Left            =   6150
         TabIndex        =   251
         Top             =   2220
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   6
            ItemData        =   "frmConfi.frx":C835
            Left            =   1230
            List            =   "frmConfi.frx":C842
            Style           =   2  'Dropdown List
            TabIndex        =   252
            ToolTipText     =   "Select Response to Air Temperature OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   240
            TabIndex        =   253
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Can Vent OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   810
         Index           =   7
         Left            =   3075
         TabIndex        =   248
         Top             =   3030
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   7
            ItemData        =   "frmConfi.frx":C85D
            Left            =   1230
            List            =   "frmConfi.frx":C86A
            Style           =   2  'Dropdown List
            TabIndex        =   249
            ToolTipText     =   "Select Response to CanVent OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   240
            TabIndex        =   250
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Butane Flow OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Index           =   1
         Left            =   0
         TabIndex        =   245
         Top             =   600
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            ItemData        =   "frmConfi.frx":C885
            Left            =   1230
            List            =   "frmConfi.frx":C892
            Style           =   2  'Dropdown List
            TabIndex        =   246
            ToolTipText     =   "Select Response to Butane Flow Rate OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   247
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Nitrogen Flow OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Index           =   2
         Left            =   0
         TabIndex        =   242
         Top             =   1410
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            ItemData        =   "frmConfi.frx":C8AD
            Left            =   1230
            List            =   "frmConfi.frx":C8BA
            Style           =   2  'Dropdown List
            TabIndex        =   243
            ToolTipText     =   "Select Response to Nitrogen (or Vapor Carrier) Flow Rate OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   240
            TabIndex        =   244
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Fuel Temp OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Index           =   3
         Left            =   0
         TabIndex        =   239
         Top             =   2220
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            ItemData        =   "frmConfi.frx":C8D5
            Left            =   1230
            List            =   "frmConfi.frx":C8E2
            Style           =   2  'Dropdown List
            TabIndex        =   240
            ToolTipText     =   "Select Response to LiveFuel Temperature OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   240
            TabIndex        =   241
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Air Moisture OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   810
         Index           =   5
         Left            =   6150
         TabIndex        =   236
         Top             =   1410
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   5
            ItemData        =   "frmConfi.frx":C8FD
            Left            =   1230
            List            =   "frmConfi.frx":C90A
            Style           =   2  'Dropdown List
            TabIndex        =   237
            ToolTipText     =   "Select Response to Air Moisture OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   240
            TabIndex        =   238
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Purge Flow OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   810
         Index           =   4
         Left            =   6120
         TabIndex        =   233
         Top             =   600
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            ItemData        =   "frmConfi.frx":C925
            Left            =   1230
            List            =   "frmConfi.frx":C932
            Style           =   2  'Dropdown List
            TabIndex        =   234
            ToolTipText     =   "Select Response to Purge Flow Rate OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   240
            TabIndex        =   235
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Load Rate OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   810
         Index           =   8
         Left            =   0
         TabIndex        =   230
         Top             =   3030
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            ItemData        =   "frmConfi.frx":C94D
            Left            =   1230
            List            =   "frmConfi.frx":C95A
            Style           =   2  'Dropdown List
            TabIndex        =   231
            ToolTipText     =   "Select Response to LoadRate PID OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   240
            TabIndex        =   232
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.Frame frmOotResp 
         Caption         =   "Purge DP OOT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   810
         Index           =   9
         Left            =   6150
         TabIndex        =   227
         Top             =   3030
         Width           =   3075
         Begin VB.ComboBox ResponseOOT 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            ItemData        =   "frmConfi.frx":C975
            Left            =   1230
            List            =   "frmConfi.frx":C982
            Style           =   2  'Dropdown List
            TabIndex        =   228
            ToolTipText     =   "Select Response to Purge Flow Rate OutOfTolerance condition"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblLimitSw 
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Response:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   240
            TabIndex        =   229
            Top             =   405
            Width           =   1125
         End
      End
      Begin VB.TextBox txtPurgeDpHiLimit 
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
         Left            =   -66600
         TabIndex        =   218
         Text            =   "4.0"
         ToolTipText     =   "-5.0 to +5.0   inches of H2O"
         Top             =   1890
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageFillShutoff 
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
         Left            =   -71445
         TabIndex        =   216
         Text            =   "0"
         ToolTipText     =   "0-100 % for Shutoff of Fill Operation"
         Top             =   2925
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageFillTimeout 
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
         Left            =   -71445
         TabIndex        =   215
         Text            =   "0"
         ToolTipText     =   "5-999 Sec for Fill Operation Timeout"
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageFillDelay 
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
         Left            =   -71445
         TabIndex        =   214
         Text            =   "0"
         ToolTipText     =   "Settle Time After Pump ShutOff (0-99)"
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageDrainShutoff 
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
         Left            =   -72465
         TabIndex        =   207
         Text            =   "0"
         ToolTipText     =   "0-100 % for Shutoff of Drain Operation"
         Top             =   2925
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageDrainTimeout 
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
         Left            =   -72465
         TabIndex        =   206
         Text            =   "0"
         ToolTipText     =   "5-999 Sec for Drain Operation Timeout"
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageDrainDelay 
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
         Left            =   -72465
         TabIndex        =   205
         Text            =   "0"
         ToolTipText     =   "Pump ShutOff Delay After Low Switch (0-99)"
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtFuelStorageTankVol 
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
         Left            =   -71445
         TabIndex        =   198
         Text            =   "199.9"
         ToolTipText     =   "Volume of the Fuel Storage Tank"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtVaporGenTankVol 
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
         Left            =   -66720
         TabIndex        =   195
         Text            =   "199.9"
         ToolTipText     =   "Volume of the Vapor Generator Tank"
         Top             =   3480
         Width           =   735
      End
      Begin VB.TextBox txtLoadRate_Pgain 
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
         Left            =   -71400
         TabIndex        =   189
         Text            =   "199.9"
         ToolTipText     =   "PID Controller Proportional Gain"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLoadRate_Igain 
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
         Left            =   -71400
         TabIndex        =   188
         Text            =   "199.9"
         ToolTipText     =   "PID Controller Integral Gain"
         Top             =   1140
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgPurgeFillDelay 
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
         Left            =   -66720
         TabIndex        =   171
         Text            =   "0"
         ToolTipText     =   "Delay for N2 Purge (after Fill) in sec"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgPurgeTimeout 
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
         Left            =   -66705
         TabIndex        =   170
         Text            =   "0"
         ToolTipText     =   "5-999 Sec for Purge PS Timeout"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgPurgeDrainDelay 
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
         Left            =   -66720
         TabIndex        =   169
         Text            =   "0"
         ToolTipText     =   "Delay for N2 Purge (before Drain) in sec"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgDrainDelay 
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
         Left            =   -67695
         TabIndex        =   168
         Text            =   "0"
         ToolTipText     =   "Pump ShutOff Delay After Low Switch (0-99) in seconds"
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgFillDelay 
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
         Left            =   -66705
         TabIndex        =   167
         Text            =   "0"
         ToolTipText     =   "Settle Time After Pump ShutOff (0-99)  in seconds"
         Top             =   2340
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgDrainTimeout 
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
         Left            =   -67695
         TabIndex        =   166
         Text            =   "0"
         ToolTipText     =   "5-999 Seconds for Drain Operation Timeout"
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgFillTimeout 
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
         Left            =   -66705
         TabIndex        =   165
         Text            =   "0"
         ToolTipText     =   "5-999 Seconds for Fill Operation Timeout"
         Top             =   2625
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgHeaterTimeout 
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
         Left            =   -66705
         TabIndex        =   164
         Text            =   "0"
         ToolTipText     =   "5-99 Min for Heat-to-Temp Operation Timeout"
         Top             =   1620
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgFillShutoff 
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
         Left            =   -66720
         TabIndex        =   163
         Text            =   "0"
         ToolTipText     =   "0-100 % for Shutoff of Fill Operation"
         Top             =   2925
         Width           =   735
      End
      Begin VB.TextBox txtLiveFuelChgDrainShutoff 
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
         Left            =   -67695
         TabIndex        =   162
         Text            =   "0"
         ToolTipText     =   "0-100 % for Shutoff of Drain Operation"
         Top             =   2925
         Width           =   735
      End
      Begin VB.Frame frmStnSelection 
         Caption         =   "LiveFuel Station"
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
         Height          =   975
         Left            =   -74880
         TabIndex        =   158
         Top             =   480
         Width           =   2325
         Begin VB.TextBox txtDispStn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
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
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   615
            Left            =   765
            MaxLength       =   2
            TabIndex        =   161
            Text            =   "9"
            Top             =   253
            Width           =   720
         End
         Begin VB.CommandButton cmdStnDn 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfi.frx":C99D
            Style           =   1  'Graphical
            TabIndex        =   160
            ToolTipText     =   "previous livefuel station"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   640
         End
         Begin VB.CommandButton cmdStnUp 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   1485
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmConfi.frx":D09F
            Style           =   1  'Graphical
            TabIndex        =   159
            ToolTipText     =   "next live fuel station"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   640
         End
      End
      Begin VB.TextBox txtOnDutyMult 
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
         Index           =   1
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   156
         Text            =   "1.00"
         ToolTipText     =   "Heater On Duty Multiplier (0.8 to 1.5)"
         Top             =   585
         Width           =   735
      End
      Begin VB.TextBox txtOffDutyMult 
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
         Index           =   1
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   155
         Text            =   "1.00"
         ToolTipText     =   "Heater Off Duty Multiplier ( (0.5 to 2.0))"
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtPgain 
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
         Index           =   2
         Left            =   -66690
         TabIndex        =   154
         Text            =   "199.9"
         ToolTipText     =   "PID Controller Proportional Gain"
         Top             =   2205
         Width           =   735
      End
      Begin VB.TextBox txtOOTtime 
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
         Left            =   -70890
         TabIndex        =   153
         Text            =   "29"
         ToolTipText     =   "Enter 1 to 999 seconds to allow MFC settle down"
         Top             =   585
         Width           =   735
      End
      Begin VB.CheckBox optLogTempRhVerbose 
         Alignment       =   1  'Right Justify
         Caption         =   "All TempRh Log events to EventLog ?"
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
         Left            =   -69600
         TabIndex        =   150
         ToolTipText     =   "Record All TempRh Log events to EventLog"
         Top             =   3510
         Width           =   3615
      End
      Begin VB.TextBox txtLogTempRhInterval 
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
         Left            =   -70785
         MaxLength       =   4
         TabIndex        =   147
         Text            =   "100"
         ToolTipText     =   "10 to 999 Minutes"
         Top             =   3480
         Width           =   735
      End
      Begin VB.Frame frmGenOptions 
         Caption         =   "Manual Report Generation Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3495
         Left            =   -70320
         TabIndex        =   140
         Top             =   480
         Width           =   4435
         Begin VB.CheckBox chkCsvGenReporting 
            Caption         =   "Csv Report"
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
            Left            =   840
            TabIndex        =   226
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CheckBox chkCsvGenSummary 
            Caption         =   "Include Summary Data"
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
            Left            =   1020
            TabIndex        =   225
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox chkCsvGenDetail 
            Caption         =   "Include Detail Data"
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
            Left            =   1020
            TabIndex        =   224
            Top             =   3000
            Width           =   2415
         End
         Begin VB.CheckBox chkTextGenSummary 
            Caption         =   "Summary Report"
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
            Left            =   1020
            TabIndex        =   146
            Top             =   720
            Width           =   1935
         End
         Begin VB.CheckBox chkTextGenReporting 
            Caption         =   "Text Reports"
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
            Left            =   840
            TabIndex        =   145
            Top             =   480
            Width           =   2055
         End
         Begin VB.CheckBox chkTextGenDetail 
            Caption         =   "Detail Report"
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
            Left            =   1020
            TabIndex        =   144
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox chkXlsGenDetail 
            Caption         =   "Include Detail Data"
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
            Left            =   1020
            TabIndex        =   143
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox chkXlsGenSummary 
            Caption         =   "Include Summary Data"
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
            Left            =   1020
            TabIndex        =   142
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CheckBox chkXlsGenReporting 
            Caption         =   "XLS Report"
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
            Left            =   840
            TabIndex        =   141
            Top             =   1560
            Width           =   2055
         End
      End
      Begin VB.Frame frmEotOptions 
         Caption         =   "End-of-Test Reporting Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3495
         Left            =   -74880
         TabIndex        =   136
         Top             =   480
         Width           =   4435
         Begin VB.CheckBox chkCsvEotDetail 
            Caption         =   "Include Detail Data"
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
            Left            =   900
            TabIndex        =   223
            Top             =   3000
            Width           =   2415
         End
         Begin VB.CheckBox chkCsvEotSummary 
            Caption         =   "Include Summary Data"
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
            Left            =   900
            TabIndex        =   222
            Top             =   2760
            Width           =   2415
         End
         Begin VB.CheckBox chkCsvEotReporting 
            Caption         =   "Csv Report"
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
            Left            =   720
            TabIndex        =   221
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CheckBox chkXlsEotReporting 
            Caption         =   "XLS Report"
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
            Left            =   840
            TabIndex        =   194
            Top             =   1560
            Width           =   2055
         End
         Begin VB.CheckBox chkXlsEotSummary 
            Caption         =   "Include Summary Data"
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
            Left            =   1020
            TabIndex        =   193
            Top             =   1800
            Width           =   2415
         End
         Begin VB.CheckBox chkXlsEotDetail 
            Caption         =   "Include Detail Data"
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
            Left            =   1020
            TabIndex        =   192
            Top             =   2040
            Width           =   2415
         End
         Begin VB.CheckBox chkTextEotDetail 
            Caption         =   "Detail Report"
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
            Left            =   1020
            TabIndex        =   151
            Top             =   1080
            Width           =   3135
         End
         Begin VB.CheckBox chkTextEotSummaryAutoPrint 
            Caption         =   "Auto-Print Summary Report"
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
            Left            =   1200
            TabIndex        =   139
            Top             =   840
            Width           =   2835
         End
         Begin VB.CheckBox chkTextEotSummary 
            Caption         =   "Summary Report"
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
            Left            =   1020
            TabIndex        =   138
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox chkTextEotReporting 
            Caption         =   "Text Reports"
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
            Left            =   840
            TabIndex        =   137
            Top             =   360
            Width           =   3495
         End
      End
      Begin VB.TextBox txtLoadTimeLimit 
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
         Left            =   -71205
         MaxLength       =   4
         TabIndex        =   133
         Text            =   "1.05"
         ToolTipText     =   "Multiplier of Recipe Load Time (1.05 - 5.0)"
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtButaneMassLimit 
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
         Left            =   -71205
         MaxLength       =   4
         TabIndex        =   130
         Text            =   "1.05"
         ToolTipText     =   "Multiplier of Canister Working Capacity (1.05 - 5.0)"
         Top             =   3000
         Width           =   735
      End
      Begin VB.ComboBox SelectUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfi.frx":D7A1
         Left            =   -68280
         List            =   "frmConfi.frx":D7A8
         Style           =   2  'Dropdown List
         TabIndex        =   128
         ToolTipText     =   "User to be Logged on at Startup"
         Top             =   1650
         Width           =   2325
      End
      Begin VB.TextBox txtPurgeSettleTime 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   126
         Text            =   "0"
         ToolTipText     =   "Time to allow scale values to settle"
         Top             =   3720
         Width           =   715
      End
      Begin VB.TextBox txtLoadSettleTime 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   124
         Text            =   "0"
         ToolTipText     =   "Time to allow scale values to settle"
         Top             =   3720
         Width           =   715
      End
      Begin VB.TextBox txtORVRBtnTol 
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
         Left            =   -66600
         TabIndex        =   118
         Text            =   "0.00"
         ToolTipText     =   "Butane Flow Tolerance in Grams per Hour (1-1999)"
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtORVRNitTol 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   117
         Text            =   "0.00"
         ToolTipText     =   "Percent of full range from .1 to 100 "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox txtTimeoutDuration 
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
         Index           =   2
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   113
         Text            =   "1.00"
         ToolTipText     =   "Number of Seconds the PAS Moisture must be Out of Tolerance for a PAS Moisture Timeout"
         Top             =   3120
         Width           =   735
      End
      Begin VB.TextBox txtTimeoutDuration 
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
         Index           =   1
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   111
         Text            =   "1.00"
         ToolTipText     =   "Number of Seconds the PAS Temperature must be Out of Tolerance for a PAS Temp Timeout"
         Top             =   1545
         Width           =   735
      End
      Begin VB.TextBox txtIgain 
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
         Index           =   2
         Left            =   -66690
         TabIndex        =   108
         Text            =   "199.9"
         ToolTipText     =   "PID Controller Integral Gain"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtInTolDuration 
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
         Index           =   2
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   106
         Text            =   "1.00"
         ToolTipText     =   "Number of Seconds the PAS Moisture must be In Tolerance before PAS Ready"
         Top             =   2835
         Width           =   735
      End
      Begin VB.TextBox txtInTolDuration 
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
         Index           =   1
         Left            =   -66690
         MaxLength       =   4
         TabIndex        =   103
         Text            =   "1.00"
         ToolTipText     =   "Number of Seconds the PAS Temperature must be In Tolerance before PAS Ready"
         Top             =   1230
         Width           =   735
      End
      Begin VB.TextBox txtJobRecs 
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
         Left            =   -70755
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "100"
         ToolTipText     =   "0 to 1000 Records"
         Top             =   930
         Width           =   735
      End
      Begin VB.TextBox txtEventRecs 
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
         Left            =   -70755
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "100"
         ToolTipText     =   "0 to 1000 Records"
         Top             =   570
         Width           =   735
      End
      Begin VB.TextBox txtDoorOpenDelay 
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
         Left            =   -70890
         TabIndex        =   39
         Text            =   "15"
         ToolTipText     =   "Max delay  from 1 to 99 minutes"
         Top             =   1590
         Width           =   735
      End
      Begin VB.TextBox txtUPSOpenDelay 
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
         Left            =   -70890
         TabIndex        =   38
         Text            =   "15"
         ToolTipText     =   "Max delay  from 1 to 99 minutes"
         Top             =   1905
         Width           =   735
      End
      Begin VB.TextBox txtPurgeTotIntvl 
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
         Left            =   -71205
         MaxLength       =   6
         TabIndex        =   37
         Text            =   "60"
         ToolTipText     =   "0.1 to 5 seconds (must NOT be greater than report interval) "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLoLimPurge 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   36
         Text            =   "0.0"
         ToolTipText     =   "0.0 to 4.0 %"
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox txtMoistureTarget 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   35
         Text            =   "35"
         ToolTipText     =   "0 to 200 Grains per Lb"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempTarget 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   34
         Text            =   "25"
         ToolTipText     =   "0 to 100 Degrees C"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtPurgeTotal 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   33
         Text            =   "0.0"
         ToolTipText     =   "0 to 100"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtMoistureTol 
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
         Left            =   -66600
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "3.0"
         ToolTipText     =   "0 to 200 Grains per Lb"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtTempTol 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "9.0"
         ToolTipText     =   "0 to 100 Degrees C"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtPurgeIntvl 
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
         Left            =   -71205
         MaxLength       =   3
         TabIndex        =   30
         Text            =   "60"
         ToolTipText     =   "1 to 900 seconds"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtPurgeTol 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   29
         Text            =   "0.00"
         ToolTipText     =   "Percent of full range from .1 to 100 "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtFileName 
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
         Left            =   -72495
         MaxLength       =   6
         TabIndex        =   28
         Text            =   "012345"
         ToolTipText     =   "0 to 999999"
         Top             =   1920
         Width           =   2440
      End
      Begin VB.CheckBox optDbfBackup 
         Caption         =   "Backup DB Files? "
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
         Left            =   -74880
         TabIndex        =   27
         ToolTipText     =   "File Backup Path"
         Top             =   2670
         Width           =   2295
      End
      Begin VB.TextBox txtDbfBackupPath 
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
         Left            =   -72495
         TabIndex        =   26
         ToolTipText     =   "Backup Path for DB Files"
         Top             =   2655
         Width           =   6585
      End
      Begin VB.ComboBox cboReportName2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfi.frx":D7BB
         Left            =   -70275
         List            =   "frmConfi.frx":D7CE
         TabIndex        =   25
         Text            =   "<second element>_"
         ToolTipText     =   "Second Part of the Report File Name"
         Top             =   1185
         Width           =   2200
      End
      Begin VB.TextBox txtRptBackupPath 
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
         Left            =   -72495
         TabIndex        =   24
         ToolTipText     =   "Backup Path for Report Files"
         Top             =   2955
         Width           =   6585
      End
      Begin VB.CheckBox optRptBackup 
         Caption         =   "Backup Reports? "
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
         Left            =   -74880
         TabIndex        =   23
         ToolTipText     =   "File Backup Path"
         Top             =   2970
         Width           =   1815
      End
      Begin VB.ComboBox cboReportName3 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfi.frx":D817
         Left            =   -68055
         List            =   "frmConfi.frx":D82A
         TabIndex        =   22
         Text            =   "<third element>_"
         ToolTipText     =   "Third Part of the Report File Name"
         Top             =   1185
         Width           =   2200
      End
      Begin VB.ComboBox cboReportName1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfi.frx":D873
         Left            =   -72495
         List            =   "frmConfi.frx":D883
         TabIndex        =   21
         Text            =   "<first element>_"
         ToolTipText     =   "First Part of the Report File Name"
         Top             =   1185
         Width           =   2200
      End
      Begin VB.TextBox txtHeading 
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
         Left            =   -72495
         MaxLength       =   60
         TabIndex        =   20
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   480
         Width           =   6640
      End
      Begin VB.TextBox txtHeading2 
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
         Left            =   -72495
         MaxLength       =   60
         TabIndex        =   19
         ToolTipText     =   "Alphanumeric Entry"
         Top             =   840
         Width           =   6640
      End
      Begin VB.TextBox txtFuelTempTol 
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
         Left            =   -66600
         MaxLength       =   4
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "2 to 15 deg F"
         Top             =   2295
         Width           =   735
      End
      Begin VB.TextBox txtLfvTol 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   17
         Text            =   "0.0"
         ToolTipText     =   "0.1 to 100% Fullscale"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtNitTol 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   16
         Text            =   "0.00"
         ToolTipText     =   "Percent of full range from .1 to 100 "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtLoadTotIntvl 
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
         Left            =   -71205
         MaxLength       =   6
         TabIndex        =   15
         Text            =   "0"
         ToolTipText     =   "0.1 to 5 seconds (must NOT be greater than report interval) "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtLoLimLoad 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "0.0"
         ToolTipText     =   "0.0 to 4.0 %"
         Top             =   2685
         Width           =   735
      End
      Begin VB.TextBox txtCanventOvr 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   13
         Text            =   "480"
         ToolTipText     =   "0 to 29999 seconds"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtLoadPressure 
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
         Left            =   -71205
         MaxLength       =   5
         TabIndex        =   12
         Text            =   "5.25"
         ToolTipText     =   "0.5  to 15 psi"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtNitrogenPurgeTime 
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
         Left            =   -71205
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "0"
         ToolTipText     =   "PreLoad N2 Purge Duration in seconds (0 to 900)"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtLoadIntvl 
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
         Left            =   -71205
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "0"
         ToolTipText     =   "1 to 900 seconds"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtLoadTotal 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0.0"
         ToolTipText     =   "0 to 100 Percent"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtMixRatio 
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
         Left            =   -66600
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0.0"
         ToolTipText     =   "0 to 100"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox txtBtnTol 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "0.00"
         ToolTipText     =   "Butane Flow Tolerance in Grams per Hour (0.2-20)"
         Top             =   840
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox SelectLkChkFailResp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmConfi.frx":D8C3
         Left            =   -68205
         List            =   "frmConfi.frx":D8D3
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select Response to Leak Check Failure"
         Top             =   480
         Width           =   2370
      End
      Begin VB.TextBox txtLCIntvl 
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
         Left            =   -71205
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "60"
         ToolTipText     =   "1 to 900 seconds"
         Top             =   495
         Width           =   735
      End
      Begin VB.TextBox txtLCTime 
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
         Left            =   -71205
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "0.00"
         ToolTipText     =   "10 to 300 sec."
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtLCMinDelay 
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
         Left            =   -71205
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "60"
         ToolTipText     =   "30 to 999 sec."
         Top             =   855
         Width           =   735
      End
      Begin VB.TextBox txtPressureDecay 
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
         Left            =   -66600
         MaxLength       =   4
         TabIndex        =   2
         Text            =   "0.0"
         ToolTipText     =   "1 to 99 Percent"
         Top             =   1215
         Width           =   735
      End
      Begin VB.TextBox txtLCSetPoint 
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
         Left            =   -66600
         MaxLength       =   5
         TabIndex        =   1
         Text            =   "15"
         ToolTipText     =   "0.1 to 15 PSI"
         Top             =   855
         Width           =   735
      End
      Begin VB.Label lblRemStatusLogIntervalUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -71700
         TabIndex        =   290
         Top             =   4215
         Width           =   855
      End
      Begin VB.Label lblRemStatusLogInterval 
         BackStyle       =   0  'Transparent
         Caption         =   "Remote Status Data Log Interval:"
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
         Left            =   -74760
         TabIndex        =   289
         Top             =   4215
         Width           =   2895
      End
      Begin VB.Label lblOvenTempTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Oven Temp Tolerance:"
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
         Left            =   -69960
         TabIndex        =   286
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lblOvenTempUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -67215
         TabIndex        =   285
         Top             =   4095
         Width           =   645
      End
      Begin VB.Label lblWaterBathUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -72000
         TabIndex        =   283
         Top             =   2295
         Width           =   735
      End
      Begin VB.Label lblWaterBathDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "WaterBath Temp Tolerance:"
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
         Left            =   -74880
         TabIndex        =   282
         Top             =   2295
         Width           =   2640
      End
      Begin VB.Label lblOvenBand 
         BackStyle       =   0  'Transparent
         Caption         =   "Oven  +/- Temperature OK:"
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
         Left            =   -69960
         TabIndex        =   271
         Top             =   3735
         Width           =   2535
      End
      Begin VB.Label lblOvenBandUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -67215
         TabIndex        =   270
         Top             =   3735
         Width           =   645
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Default Report Interval:"
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
         Left            =   -74850
         TabIndex        =   268
         Top             =   1095
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -71815
         TabIndex        =   267
         Top             =   1095
         Width           =   855
      End
      Begin VB.Label lblFuelStorageLeakRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Storage Level Tolerance:"
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
         Left            =   -74760
         TabIndex        =   262
         Top             =   3735
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Label lblFuelStorageLeakRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "(%FS)    "
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
         Left            =   -72000
         TabIndex        =   261
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   3735
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblVaporGenLeakRate 
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Generator Level Tolerance:"
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
         Left            =   -70440
         TabIndex        =   259
         Top             =   3720
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblVaporGenLeakRateUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "(%FS)   "
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
         Left            =   -67320
         TabIndex        =   258
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   3735
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label lblPurgeDpHiLimit 
         BackStyle       =   0  'Transparent
         Caption         =   "Diffential Pressure Hi Limit:"
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
         Left            =   -70005
         TabIndex        =   220
         Top             =   1905
         Width           =   2655
      End
      Begin VB.Label lblPurgeDpHiLimitUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(in H2O)"
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
         Left            =   -67500
         TabIndex        =   219
         Top             =   1905
         Width           =   885
      End
      Begin VB.Label lblLoadByPID 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Load By PID"
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
         Height          =   255
         Left            =   -72360
         TabIndex        =   217
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label lblFST_Shutoff2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% level)"
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
         Left            =   -73785
         TabIndex        =   213
         Top             =   2955
         Width           =   1005
      End
      Begin VB.Label lblFST_Shutoff 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutoff:"
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
         Left            =   -74760
         TabIndex        =   212
         Top             =   2955
         Width           =   900
      End
      Begin VB.Label lblFST_Timeout 
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout:"
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
         Left            =   -74760
         TabIndex        =   211
         Top             =   2655
         Width           =   900
      End
      Begin VB.Label lblFST_Timeout2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -73785
         TabIndex        =   210
         Top             =   2655
         Width           =   1005
      End
      Begin VB.Label lblFST_Delay 
         BackStyle       =   0  'Transparent
         Caption         =   "Delay:"
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
         Left            =   -74760
         TabIndex        =   209
         Top             =   2355
         Width           =   900
      End
      Begin VB.Label lblFST_Delay2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -73785
         TabIndex        =   208
         Top             =   2355
         Width           =   1005
      End
      Begin VB.Label lblStorageTank 
         BackStyle       =   0  'Transparent
         Caption         =   "Storage Tank"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   204
         Top             =   2100
         Width           =   1605
      End
      Begin VB.Label lblFST_Fill 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fill"
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
         Height          =   255
         Left            =   -71445
         TabIndex        =   203
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblFST_Drain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drain"
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
         Height          =   255
         Left            =   -72465
         TabIndex        =   202
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblVaportank 
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Tank "
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
         Height          =   255
         Left            =   -69990
         TabIndex        =   201
         Top             =   2100
         Width           =   1605
      End
      Begin VB.Label lblFuelStorageTankVol2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(gallons) "
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
         Left            =   -72390
         TabIndex        =   200
         Top             =   3510
         Width           =   885
      End
      Begin VB.Label lblFuelStorageTankVol 
         BackStyle       =   0  'Transparent
         Caption         =   "Fuel Storage Tank Volume"
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
         Left            =   -74760
         TabIndex        =   199
         Top             =   3510
         Width           =   2385
      End
      Begin VB.Label lblVaporGenTankVol2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(gallons) "
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
         Left            =   -67680
         TabIndex        =   197
         Top             =   3510
         Width           =   885
      End
      Begin VB.Label lblVaporGenTankVol 
         BackStyle       =   0  'Transparent
         Caption         =   "Vapor Generator Tank Volume:"
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
         Left            =   -70350
         TabIndex        =   196
         Top             =   3510
         Width           =   2640
      End
      Begin VB.Label lblLoadRate_Pgain 
         BackStyle       =   0  'Transparent
         Caption         =   "PID - P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   191
         Top             =   855
         Width           =   840
      End
      Begin VB.Label lblLoadRate_Igain 
         BackStyle       =   0  'Transparent
         Caption         =   "PID - I "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72360
         TabIndex        =   190
         Top             =   1155
         Width           =   840
      End
      Begin VB.Label lblADF_PurgeFillDelay2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -67800
         TabIndex        =   187
         Top             =   930
         Width           =   1005
      End
      Begin VB.Label lblADF_PurgeFillDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Pressurize Delay after Fill:"
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
         Left            =   -70350
         TabIndex        =   186
         Top             =   930
         Width           =   2700
      End
      Begin VB.Label lblADF_PurgeTimeout 
         BackStyle       =   0  'Transparent
         Caption         =   "Pressurize Timeout:"
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
         Left            =   -70350
         TabIndex        =   185
         Top             =   1230
         Width           =   2700
      End
      Begin VB.Label lblADF_PurgeTimeout2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -67800
         TabIndex        =   184
         Top             =   1230
         Width           =   1005
      End
      Begin VB.Label lblADF_PurgeDrainDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Pressurize Delay before Drain:"
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
         Left            =   -70350
         TabIndex        =   183
         Top             =   630
         Width           =   2700
      End
      Begin VB.Label lblADF_PurgeDrainDelay2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -67800
         TabIndex        =   182
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblADF_Drain 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Drain"
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
         Height          =   255
         Left            =   -67695
         TabIndex        =   181
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblADF_Fill 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fill"
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
         Height          =   255
         Left            =   -66675
         TabIndex        =   180
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label lblADF_Delay2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -69015
         TabIndex        =   179
         Top             =   2355
         Width           =   1005
      End
      Begin VB.Label lblADF_Delay 
         BackStyle       =   0  'Transparent
         Caption         =   "Delay:"
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
         Left            =   -69990
         TabIndex        =   178
         Top             =   2355
         Width           =   900
      End
      Begin VB.Label lblADF_Timeout2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -69015
         TabIndex        =   177
         Top             =   2655
         Width           =   1005
      End
      Begin VB.Label lblADF_Timeout 
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout:"
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
         Left            =   -69990
         TabIndex        =   176
         Top             =   2655
         Width           =   900
      End
      Begin VB.Label lblADF_HeaterTimeout 
         BackStyle       =   0  'Transparent
         Caption         =   "Heater Timeout:"
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
         Left            =   -70350
         TabIndex        =   175
         Top             =   1650
         Width           =   1455
      End
      Begin VB.Label lblADF_HeaterTimeout2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(minutes)"
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
         Left            =   -67695
         TabIndex        =   174
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label lblADF_Shutoff 
         BackStyle       =   0  'Transparent
         Caption         =   "Shutoff:"
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
         Left            =   -69990
         TabIndex        =   173
         Top             =   2955
         Width           =   900
      End
      Begin VB.Label lblADF_Shutoff2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% level)"
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
         Left            =   -69015
         TabIndex        =   172
         Top             =   2955
         Width           =   1005
      End
      Begin VB.Label Label40 
         BackStyle       =   0  'Transparent
         Caption         =   "OOT Delay in Seconds:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -74850
         TabIndex        =   152
         Top             =   600
         Width           =   3930
      End
      Begin VB.Label lblLogTempRhIntervalUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutes:"
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
         Left            =   -71595
         TabIndex        =   149
         Top             =   3510
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label lblLogTempRhInterval 
         BackStyle       =   0  'Transparent
         Caption         =   "TempRh Log Interval"
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
         Left            =   -74760
         TabIndex        =   148
         ToolTipText     =   "Minutes between successive Temp & Rh Log Records"
         Top             =   3510
         Width           =   2775
      End
      Begin VB.Label lblLoadTimeLimit 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Time Limit:                     "
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
         Left            =   -74880
         TabIndex        =   135
         ToolTipText     =   "Enter a multiplier of the canister's working capacity ("
         Top             =   3375
         Width           =   1920
      End
      Begin VB.Label lblLoadTimeLimitUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(LoadTime mult)"
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
         Left            =   -72690
         TabIndex        =   134
         Top             =   3375
         Width           =   1425
      End
      Begin VB.Label lblButaneMassLimit 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane Mass Limit:                     "
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
         Left            =   -74880
         TabIndex        =   132
         ToolTipText     =   "Enter a multiplier of the canister's working capacity ("
         Top             =   3015
         Width           =   2160
      End
      Begin VB.Label lblButaneMassLimitUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(WC mult)"
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
         Left            =   -72480
         TabIndex        =   131
         Top             =   3015
         Width           =   1185
      End
      Begin VB.Label lblAutoLogon 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Logon"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69540
         TabIndex        =   129
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblPurgeSettleTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Settling Time                   (minutes)"
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
         Left            =   -74880
         TabIndex        =   127
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label lblLoadSettleTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Settling Time                    (minutes)"
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
         Left            =   -74880
         TabIndex        =   125
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label lblORVRButFlowUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Gms/Hr)"
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
         Left            =   -67770
         TabIndex        =   120
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   3375
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblORVRNitFlowUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67800
         TabIndex        =   119
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   3735
         Width           =   1110
      End
      Begin VB.Label lblORVRButFlowTol 
         BackStyle       =   0  'Transparent
         Caption         =   "ORVR But Flow Tolerance:"
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
         Left            =   -70080
         TabIndex        =   116
         Top             =   3360
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblORVRNitFlowTol 
         BackStyle       =   0  'Transparent
         Caption         =   "ORVR N2 Flow Tolerance:"
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
         Left            =   -70080
         TabIndex        =   115
         Top             =   3720
         Width           =   2280
      End
      Begin VB.Label lblTimeoutDuration 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout Duration in Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -69255
         TabIndex        =   114
         Top             =   3150
         Width           =   2520
      End
      Begin VB.Label lblTimeoutDuration 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Timeout Duration in Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -69255
         TabIndex        =   112
         Top             =   1560
         Width           =   2520
      End
      Begin VB.Label lblIgain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -67080
         TabIndex        =   110
         Top             =   2535
         Width           =   240
      End
      Begin VB.Label lblPgain 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -67080
         TabIndex        =   109
         Top             =   2220
         Width           =   240
      End
      Begin VB.Label lblInTolDuration 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "InTolerance Duration in Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -69600
         TabIndex        =   107
         Top             =   2850
         Width           =   2850
      End
      Begin VB.Label lblPidControl 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PAS Moisture Control:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   2
         Left            =   -69720
         TabIndex        =   105
         Top             =   2220
         Width           =   2100
      End
      Begin VB.Label lblInTolDuration 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "InTolerance Duration in Seconds"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -69600
         TabIndex        =   104
         Top             =   1260
         Width           =   2850
      End
      Begin VB.Label lblJobRecs 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Job List Records:"
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
         Left            =   -74730
         TabIndex        =   101
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblEventRecs 
         BackStyle       =   0  'Transparent
         Caption         =   "Max Event Log Records:"
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
         Left            =   -74730
         TabIndex        =   100
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblDoorOpenDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "Door Open Delay in Minutes:"
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
         Left            =   -74850
         TabIndex        =   99
         Top             =   1620
         Width           =   3855
      End
      Begin VB.Label lblUPSOpenDelay 
         BackStyle       =   0  'Transparent
         Caption         =   "UPS Power Down in Minutes:"
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
         Left            =   -74850
         TabIndex        =   98
         Top             =   1935
         Width           =   3855
      End
      Begin VB.Label lblPidControl 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PAS Temp Control:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Index           =   1
         Left            =   -69705
         TabIndex        =   97
         Top             =   615
         Width           =   1740
      End
      Begin VB.Label lblOffDutyMult 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OffDuty Mult"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -67845
         TabIndex        =   96
         Top             =   930
         Width           =   1200
      End
      Begin VB.Label lblOnDutyMult 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OnDuty Mult"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -67845
         TabIndex        =   95
         Top             =   615
         Width           =   1200
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Log Data Interval:"
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
         Left            =   -74880
         TabIndex        =   94
         Top             =   855
         Width           =   2400
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72090
         TabIndex        =   93
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Limit to Purge Flow Tolerance Checking:"
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
         Left            =   -71700
         TabIndex        =   92
         Top             =   2220
         Width           =   3900
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67725
         TabIndex        =   91
         ToolTipText     =   "Low Limit as a percent of full scale for Station MFC"
         Top             =   2220
         Width           =   1110
      End
      Begin VB.Label lblMoistTargetUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Grns/Lb)"
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
         Left            =   -72210
         TabIndex        =   90
         Top             =   1575
         Width           =   975
      End
      Begin VB.Label lblMoistureTarget 
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture Target:"
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
         Left            =   -74880
         TabIndex        =   89
         Top             =   1575
         Width           =   2175
      End
      Begin VB.Label lblTempTargetUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -71880
         TabIndex        =   88
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label lblTempTarget 
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature Target:"
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
         Left            =   -74880
         TabIndex        =   87
         Top             =   1215
         Width           =   2175
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% target)"
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
         Left            =   -67470
         TabIndex        =   86
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Volume Tolerance:"
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
         Left            =   -70005
         TabIndex        =   85
         Top             =   495
         Width           =   2175
      End
      Begin VB.Label lblMoistTolUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Grns/Lb)"
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
         Left            =   -67500
         TabIndex        =   84
         Top             =   1575
         Width           =   885
      End
      Begin VB.Label lblTempTolUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -67260
         TabIndex        =   83
         Top             =   1215
         Width           =   645
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72090
         TabIndex        =   82
         Top             =   495
         Width           =   855
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67725
         TabIndex        =   81
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   855
         Width           =   1110
      End
      Begin VB.Label lblMoistureTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Moisture Tolerance:"
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
         Left            =   -70005
         TabIndex        =   80
         Top             =   1575
         Width           =   2295
      End
      Begin VB.Label lblTempTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Temperature Tolerance:"
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
         Left            =   -70005
         TabIndex        =   79
         Top             =   1215
         Width           =   2415
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Flow Tolerance:"
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
         Left            =   -70005
         TabIndex        =   78
         Top             =   855
         Width           =   1935
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Purge Report Interval:"
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
         Left            =   -74880
         TabIndex        =   77
         Top             =   495
         Width           =   1935
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Next Job #:"
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
         Left            =   -74880
         TabIndex        =   76
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label Label35 
         BackStyle       =   0  'Transparent
         Caption         =   "Report File Name:"
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
         Left            =   -74880
         TabIndex        =   75
         Top             =   1215
         Width           =   2295
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Heading Line 1: "
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
         Left            =   -74880
         TabIndex        =   74
         Top             =   495
         Width           =   2295
      End
      Begin VB.Label Label38 
         BackStyle       =   0  'Transparent
         Caption         =   "Report Heading Line 2:"
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
         Left            =   -74880
         TabIndex        =   73
         Top             =   855
         Width           =   2295
      End
      Begin VB.Label lblFuelTempTol 
         BackStyle       =   0  'Transparent
         Caption         =   "LiveFuel Temp Tolerance:"
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
         Left            =   -70080
         TabIndex        =   72
         Top             =   2310
         Width           =   2295
      End
      Begin VB.Label lblFuelTempUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(deg C)"
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
         Left            =   -67560
         TabIndex        =   71
         Top             =   2310
         Width           =   855
      End
      Begin VB.Label lblFuelFlowUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67800
         TabIndex        =   70
         Top             =   1935
         Width           =   1095
      End
      Begin VB.Label lblFuelFlowTol 
         BackStyle       =   0  'Transparent
         Caption         =   "VaporCarrier Flow Tol.:"
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
         Left            =   -70080
         TabIndex        =   69
         Top             =   1935
         Width           =   2100
      End
      Begin VB.Label lblNitFlowUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67800
         TabIndex        =   68
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   1575
         Width           =   1110
      End
      Begin VB.Label lblNitFlowTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Nitrogen Flow Tolerance:"
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
         Left            =   -70080
         TabIndex        =   67
         Top             =   1575
         Width           =   2160
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72120
         TabIndex        =   66
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Log Data Interval:"
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
         Left            =   -74880
         TabIndex        =   65
         Top             =   855
         Width           =   2055
      End
      Begin VB.Label Label45 
         BackStyle       =   0  'Transparent
         Caption         =   "Low Limit to Load Flow Tolerance Checking:"
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
         Left            =   -71700
         TabIndex        =   64
         Top             =   2700
         Width           =   3900
      End
      Begin VB.Label Label44 
         BackStyle       =   0  'Transparent
         Caption         =   "(% full scale)"
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
         Left            =   -67785
         TabIndex        =   63
         ToolTipText     =   "Low Limit as a percent of full scale for Station MFC"
         Top             =   2700
         Width           =   1110
      End
      Begin VB.Label lblCanventUnits 
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72120
         TabIndex        =   62
         Top             =   1935
         Width           =   855
      End
      Begin VB.Label lblCanventDescr 
         BackStyle       =   0  'Transparent
         Caption         =   "Canvent FS Override Delay:"
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
         Left            =   -74880
         TabIndex        =   61
         Top             =   1935
         Width           =   2400
      End
      Begin VB.Label lblButFlowUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(Gms/Hr)"
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
         Left            =   -67770
         TabIndex        =   60
         ToolTipText     =   "Tolerance as a percent of full scale for Station MFC"
         Top             =   855
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label lblLoadPressureUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(psi)"
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
         Left            =   -71730
         TabIndex        =   59
         Top             =   1215
         Width           =   465
      End
      Begin VB.Label lblLoadPressure 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Pressure Limit:                     "
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
         Left            =   -74880
         TabIndex        =   58
         Top             =   1215
         Width           =   2160
      End
      Begin VB.Label lblNitrogenPurgeTimeUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72120
         TabIndex        =   57
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label lblNitrogenPurgeTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Nitrogen Purge Time:"
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
         Left            =   -74880
         TabIndex        =   56
         Top             =   1575
         Width           =   2055
      End
      Begin VB.Label Label53 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Report Interval:"
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
         Left            =   -74880
         TabIndex        =   55
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblButFlowTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Butane Flow Tolerance:"
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
         Left            =   -70080
         TabIndex        =   54
         Top             =   855
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72120
         TabIndex        =   53
         Top             =   495
         Width           =   855
      End
      Begin VB.Label lblMixTol 
         BackStyle       =   0  'Transparent
         Caption         =   "Mix Ratio Tolerance:"
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
         Left            =   -70080
         TabIndex        =   52
         Top             =   1215
         Width           =   1935
      End
      Begin VB.Label Label55 
         BackStyle       =   0  'Transparent
         Caption         =   "Load Total Tolerance:"
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
         Left            =   -70080
         TabIndex        =   51
         Top             =   495
         Width           =   2055
      End
      Begin VB.Label lblMixUnits 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(abs %)"
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
         Left            =   -67425
         TabIndex        =   50
         Top             =   1215
         Width           =   735
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "(% target)"
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
         Left            =   -67545
         TabIndex        =   49
         Top             =   495
         Width           =   855
      End
      Begin VB.Label lblLeakErrResponse 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check Failure:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70005
         TabIndex        =   48
         Top             =   510
         Width           =   1845
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check Report Interval:"
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
         Left            =   -74880
         TabIndex        =   47
         Top             =   510
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "(seconds)"
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
         Left            =   -72090
         TabIndex        =   46
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "Percent Allowable Pressure Decay:  (%)"
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
         Left            =   -70005
         TabIndex        =   45
         Top             =   1230
         Width           =   3375
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check Time:                           (sec)"
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
         Left            =   -74880
         TabIndex        =   44
         Top             =   1230
         Width           =   3645
      End
      Begin VB.Label lblLCSetPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check Pressure Set Point:    (PSI)"
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
         Left            =   -70005
         TabIndex        =   43
         Top             =   870
         Width           =   3390
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Leak Check Min Fill Delay:               (sec)"
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   870
         Width           =   3645
      End
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      Caption         =   "Must be showing OOT Response tab when you save the program"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2520
      TabIndex        =   274
      Top             =   8010
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   2955
      Left            =   0
      TabIndex        =   102
      Top             =   4800
      Width           =   9315
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'error mod 53 ''''''''''''''' Form CONFIG.frm ''''''''''''''''''''
Option Explicit
Private Const Tab_LeakCheck = 0
Private Const Tab_Load = 1
Private Const Tab_Purge = 2
Private Const Tab_Job = 3
Private Const Tab_Process = 4
Private Const Tab_System = 5
Private Const Tab_Reporting = 6
Private Const Tab_OotResponse = 7
Private Const Tab_AutoDrainFill = 8
Private Const SELECTRPTPATH = 0
Private Const SELECTDBFPATH = 1
Private xCol1, xCol2, xCol3, xCol4, xCol5, xCol6 As Integer
Private OotCol1Left As Integer
Private OotCol2Left As Integer
Private OotCol3Left As Integer
Private LiveFuelStn As Integer
Private NumberOfLiveFuelStations As Integer

Function Check_Config() As Boolean

' Function Name:    Check_Config
' Author:           Analytical Process Programmer     7/25/96
' Description:      Checks the validity of entries into the configuration
'                   file.  Used before saving the configuration file.
'                   Returns a true value if values are okay.
'                   Returns a false value if values are not okay.
'                   If an error is detected, an appropriate message
'                   is displayed.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 53, 1
Dim Message As String
Dim hilim As Single
Dim flag As Boolean

Check_Config = True

If Not Range_Check(txtFileName, 0, 999999, "File Name") Then Check_Config = False
If Not Range_Check(txtLCIntvl, MinDataLogSeconds, 900, "LeakCheck Report Interval") Then Check_Config = False
If Not Range_Check(txtLCMinDelay, 30, 999, "Leak Check Min Fill Delay") Then Check_Config = False
' Leak Check Test Phase Duration
If Not Range_Check(txtLCTime, 10, 300, "LeakCheck Time") Then Check_Config = False
' Leak Check Set Point
If Not Range_Check(txtLCSetPoint, 0.1, 15, "Leak Check Set Error") Then Check_Config = False
If Not Range_Check(txtPressureDecay, 1, 99, "Leak Check Pressure Decay Error") Then Check_Config = False
' check mix ratio tolerance
If Not Range_Check(txtMixRatio, 0, 100, "Mix Ratio Tol") Then Check_Config = False
' check Load Pressure Limit
If Not Range_Check(txtLoadPressure, 0.5, 15, "Load Pressure Limit") Then Check_Config = False
' check Butane Mass Limit
If USINGBUTANEMASSLIMIT Then
    If Not Range_Check(txtButaneMassLimit, 1.05, 5#, "Butane Mass Limit") Then Check_Config = False
End If
' check Load Time Limit
If USINGLOADTIMELIMIT Then
    If Not Range_Check(txtLoadTimeLimit, 1.05, 5#, "Load Time Limit") Then Check_Config = False
End If
' check Nitrogen Purge Time
If Not Range_Check(txtNitrogenPurgeTime, 0, 900, "Nitrogen Purge Time") Then Check_Config = False
' check purge total flow tolerance
If Not Range_Check(txtPurgeTotal, 0, 100, "Purge Total Tol") Then Check_Config = False
' check load total flow tolerance
If Not Range_Check(txtLoadTotal, 0, 100, "Load Total Tol") Then Check_Config = False
' check nitrogen flow tolerance
If Not Range_Check(txtNitTol, 0.1, 100, "Nitrogen Flow Tol") Then Check_Config = False
' check orvr nitrogen flow tolerance
If Not Range_Check(txtORVRNitTol, 0.1, 100, "ORVR Nitrogen Flow Tol") Then Check_Config = False
' check butane flow tolerance
If Not Range_Check(txtBtnTol, 0.2, 20, "Butane Flow Tol") Then Check_Config = False
' check orvr butane flow tolerance
If Not Range_Check(txtORVRBtnTol, 1, 1999, "ORVR Butane Flow Tol") Then Check_Config = False
' check purge flow tolerance
If Not Range_Check(txtPurgeTol, 0.1, 100, "Purge Flow Tol") Then Check_Config = False
' check load flow tolerance low limit
If Not Range_Check(txtLoLimLoad, 0#, 4#, "Load Flow Tol LowLim") Then Check_Config = False
' check purge flow tolerance low limit
If Not Range_Check(txtLoLimPurge, 0#, 4#, "Purge Flow Tol LowLim") Then Check_Config = False
' check default report interval value
If Not Range_Check(txtDefaultIntvl, MinDataLogSeconds, 900, "Default Report Interval") Then Check_Config = False
' check leakcheck report interval value
If Not Range_Check(txtLCIntvl, MinDataLogSeconds, 900, "Leakcheck Report Interval") Then Check_Config = False
' check load report interval value
If Not Range_Check(txtLoadIntvl, MinDataLogSeconds, 900, "Load Report Interval") Then Check_Config = False
' check purge report interval value
If Not Range_Check(txtPurgeIntvl, MinDataLogSeconds, 900, "Purge Report Interval") Then Check_Config = False
' check Purge Oven OK Band
If USINGPURGEOVEN Then
    If USINGC Then
        If Not Range_Check(txtOvenBand, 2#, 5#, "Purge Oven OK Band") Then Check_Config = False
        If Not Range_Check(txtOvenTempTol, 0, 100, "Purge Oven Tolerance") Then Check_Config = False
    ElseIf USINGF Then
        If Not Range_Check(txtOvenBand, 3.6, 8#, "Purge Oven OK Band") Then Check_Config = False
        If Not Range_Check(txtOvenTempTol, 0, 100, "Purge Oven Tolerance") Then Check_Config = False
    End If
End If
' Only Change PosPressPurge or DyAirPurge if AllStationsIdle
If (Not AllStationsIdle) Then
    flag = IIf((chkPosPressPurge.Value = cYES), True, False)
    If (flag <> SysConfig.PosPressPurge) Then
        Check_Config = False
        chkPosPressPurge.BackColor = PALEYELLOW
        lblMessage.Caption = lblMessage.Caption & "All Stations Must Be Idle to Change Pos Pressure Purge!" & vbCrLf
    End If
    flag = IIf((chkDryAirPurge.Value = cYES), True, False)
    If (flag <> SysConfig.DryAirPurge) Then
        Check_Config = False
        chkDryAirPurge.BackColor = PALEYELLOW
        lblMessage.Caption = lblMessage.Caption & "All Stations Must Be Idle to Change Dry Air Purge!" & vbCrLf
    End If
End If
' check load totalize interval value (must not be greater than Load Report Interval)
hilim = 10
If (hilim > ValueFromText(txtLoadIntvl.text)) Then hilim = ValueFromText(txtLoadIntvl.text)
If Not Range_Check(txtLoadTotIntvl, MinDataLogSeconds, hilim, "Load Log-Data Interval") Then Check_Config = False
' check purge totalize interval value (must not be greater than Purge Report Interval)
hilim = 10
If (hilim > ValueFromText(txtPurgeIntvl.text)) Then hilim = ValueFromText(txtPurgeIntvl.text)
If Not Range_Check(txtPurgeTotIntvl, MinDataLogSeconds, hilim, "Purge Log-Data Interval") Then Check_Config = False
If USINGC Then
    If USINGHIGHTEMPPAS Then
        ' check temperature target
        If Not Range_Check(txtTempTarget, 5, 60, "Target Temperature") Then Check_Config = False
        ' check temperature tolerance
        If Not Range_Check(txtTempTol, 0, 100, "Temperature Tol") Then Check_Config = False
    Else
        ' check temperature target
        If Not Range_Check(txtTempTarget, 5, 35, "Target Temperature") Then Check_Config = False
        ' check temperature tolerance
        If Not Range_Check(txtTempTol, 0, 100, "Temperature Tol") Then Check_Config = False
    End If
ElseIf USINGF Then
    If USINGHIGHTEMPPAS Then
        ' check temperature target
        If Not Range_Check(txtTempTarget, 40, 140, "Target Temperature") Then Check_Config = False
        ' check temperature tolerance
        If Not Range_Check(txtTempTol, 0, 100, "Temperature Tol") Then Check_Config = False
    Else
        ' check temperature target
        If Not Range_Check(txtTempTarget, 40, 90, "Target Temperature") Then Check_Config = False
        ' check temperature tolerance
        If Not Range_Check(txtTempTol, 0, 100, "Temperature Tol") Then Check_Config = False
    End If
End If
If USINGMoist_RH Then
    ' check moisture target moisture
    If Not Range_Check(txtMoistureTarget, 0, 100, "Moisture Target") Then Check_Config = False
    ' check moisture tolerance
    If Not Range_Check(txtMoistureTol, 0, 100, "Moisture Tolerance") Then Check_Config = False
ElseIf USINGMoist_Grains Then
    ' check moisture target
    If Not Range_Check(txtMoistureTarget, 0, 200, "Moisture Target") Then Check_Config = False
    ' check moisture tolerance
    If Not Range_Check(txtMoistureTol, 0, 100, "Moisture Tolerance") Then Check_Config = False
End If
' check Temp/Rh Logging values
If LogTempRh Then
    ' check log interval
    If Not Range_Check(txtLogTempRhInterval, 10#, 999#, "TempRh Log Interval") Then Check_Config = False
End If
' check Purge Differential Pressure values
If USINGPURGEDP Then
    ' check purge dp hi limit
    If Not Range_Check(txtPurgeDpHiLimit, -5, 5, "Purge DP Hi Limit") Then Check_Config = False
End If
' check Remote Status Monitor Log Interval
If USINGREMSTSMON Then
    If Not Range_Check(txtRemStatusLogInterval, 10, 60, "Remote Status Update interval") Then Check_Config = False
End If

' check Load Settle Time
If ValueFromText(txtLoadSettleTime) = 0# Then txtLoadSettleTime.text = "0"
If Not Range_Check(txtLoadSettleTime, 0#, 99.9, "Load Settle Time") Then Check_Config = False
If Not Range_Check(txtMoistureTarget, 0, 200, "Moisture Target") Then Check_Config = False
' check Purge Settle Time
If ValueFromText(txtPurgeSettleTime) = 0# Then txtPurgeSettleTime.text = "0"
If Not Range_Check(txtPurgeSettleTime, 0#, 99.9, "Purge Settle Time") Then Check_Config = False
If Not Range_Check(txtMoistureTarget, 0, 200, "Moisture Target") Then Check_Config = False
' check LiveFuel values
If systemhasLIVEFUEL Then
    ' check live fuel flow tolerance
    If Not Range_Check(txtLfvTol, 0.1, 100, "Vapor Carrier Flow Tol") Then Check_Config = False
    If systemhasADF_HEATER Then
        ' check live fuel temperature tolerance
        If USINGC Then
            If Not Range_Check(txtFuelTempTol, 1, 10, "Live Fuel Temp Tol") Then Check_Config = False
        End If
        If USINGF Then
            If Not Range_Check(txtFuelTempTol, 2, 15, "Live Fuel Temp Tol") Then Check_Config = False
        End If
    End If
    ' LiveFuel AutoDrainFill
    If ((STN_INFO(LiveFuelStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(LiveFuelStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(LiveFuelStn).Type = STN_LIVEORVR2_TYPE)) And STN_INFO(LiveFuelStn).ADF_TANKTYPE > 0 Then
        ' ADF Purge Delay before drain
        If Not Range_Check(txtLiveFuelChgPurgeDrainDelay, 0, 99, "Pressurize Delay before Drain") Then Check_Config = False
        ' ADF Purge Delay after fill
        If Not Range_Check(txtLiveFuelChgPurgeFillDelay, 0, 99, "Pressurize Delay after Fill") Then Check_Config = False
        ' ADF Purge Timeout
        If Not Range_Check(txtLiveFuelChgPurgeTimeout, 5, 99, "Pressurize Timeout") Then Check_Config = False
        ' ADF Drain Delay
        If Not Range_Check(txtLiveFuelChgDrainDelay, 0, 99, "Drain Delay") Then Check_Config = False
        ' ADF Drain Timeout
        If Not Range_Check(txtLiveFuelChgDrainTimeout, 5, 999, "Drain Timeout") Then Check_Config = False
        ' ADF Drain Shutoff
        If Not Range_Check(txtLiveFuelChgDrainShutoff, 0, 99, "Drain Shutoff") Then Check_Config = False
        ' ADF Fill Delay
        If Not Range_Check(txtLiveFuelChgFillDelay, 0, 99, "Fill Delay") Then Check_Config = False
        ' ADF Fill Timeout
        If Not Range_Check(txtLiveFuelChgFillTimeout, 5, 999, "Fill Timeout") Then Check_Config = False
        ' ADF Fill Shutoff
        If Not Range_Check(txtLiveFuelChgFillShutoff, 0, 99, "Fill Shutoff") Then Check_Config = False
        ' ADF Heater
        If systemhasADF_HEATER And STN_INFO(LiveFuelStn).ADF_TANKTYPE = 12 Then
            ' ADF Heater Timeout
            If Not Range_Check(txtLiveFuelChgHeaterTimeout, 5, 99, "Heater Timeout") Then Check_Config = False
        End If
        ' ADF Generator Tank Volume     (EU is set by LineVol Units; SI => liters, English => gallons)
        If Not Range_Check(txtVaporGenTankVol, 0.1, 199, "Vapor Tank Volume") Then Check_Config = False
        ' ADF Generator Tank Leak Rate  (EU is set by LineVol Units; SI => liters, English => gallons)
        If Not Range_Check(txtVaporGenLeakRate, 0.01, 10, "Vapor Tank Level Tolerance") Then Check_Config = False
        ' ADF Storage Tank
        If ((STN_INFO(LiveFuelStn).ADF_TANKTYPE > 20) And (STN_INFO(LiveFuelStn).ADF_TANKTYPE < 90)) Then
            ' Fuel Storage Tank Volume   (EU is set by LineVol Units; SI => liters, English => gallons)
            If Not Range_Check(txtFuelStorageTankVol, 1, 1999, "Fuel Storage Tank Volume") Then Check_Config = False
            ' Fuel Storage Tank Volume   (EU is set by LineVol Units; SI => liters, English => gallons)
            If Not Range_Check(txtFuelStorageLeakRate, 0.01, 10, "Fuel Storage Tank Level Tolerance") Then Check_Config = False
            ' FST Drain Delay
            If Not Range_Check(txtFuelStorageDrainDelay, 0, 99, "Storage Tank Drain Delay") Then Check_Config = False
            ' FST Drain Timeout
            If Not Range_Check(txtFuelStorageDrainTimeout, 5, 3600, "Storage Tank Drain Timeout") Then Check_Config = False
            ' FST Drain Shutoff
            If Not Range_Check(txtFuelStorageDrainShutoff, 0, 99, "Storage Tank Drain Shutoff") Then Check_Config = False
            ' FST Fill Delay
            If Not Range_Check(txtFuelStorageFillDelay, 0, 99, "Storage Tank Fill Delay") Then Check_Config = False
            ' FST Fill Timeout
            If Not Range_Check(txtFuelStorageFillTimeout, 5, 1200, "Storage Tank Fill Timeout") Then Check_Config = False
            ' FST Fill Shutoff
            If Not Range_Check(txtFuelStorageFillShutoff, 0, 99, "Storage Tank Fill Shutoff") Then Check_Config = False
        End If
    End If
End If

' door delay
If Not Range_Check(txtDoorOpenDelay, 1, 99, "Door Open Delay") Then Check_Config = False
If USINGUPS > 0 Then
   If USINGUPS = 1 Then
     If Not Range_Check(txtUPSOpenDelay, 0, 99, "UPS Delay") Then Check_Config = False
   End If
   If USINGUPS = 2 Then
      If Not Range_Check(txtUPSOpenDelay, 0, 10, "UPS Delay") Then Check_Config = False
   End If
Else
  If txtUPSOpenDelay = Empty Then txtUPSOpenDelay = 15
End If
If Not Range_Check(txtUPSOpenDelay, 0, 99, "UPS Delay") Then Check_Config = False
' OOT time delay
If Not Range_Check(txtOOTtime, 1, 999, "Out of Tolerance Delay") Then Check_Config = False
' Canvent Flowswitch Override time delay
If Not Range_Check(txtCanventOvr, 0, 29999, "Canvent Override Delay") Then Check_Config = False

If Not Range_Check(txtEventRecs, 0, 1000, "Event Log Records") Then Check_Config = False

If Not Range_Check(txtJobRecs, 0, 1000, "Job List Records") Then Check_Config = False

' PAS Local Control
If Not Range_Check(txtOffDutyMult(pasTEMPERATURE), 0.5, 2#, "Heater Off Duty Multiplier") Then Check_Config = False
If Not Range_Check(txtOnDutyMult(pasTEMPERATURE), 0.5, 2#, "Heater On Duty Multiplier") Then Check_Config = False
If Not Range_Check(txtInTolDuration(pasTEMPERATURE), 1, 900#, "PAS Temp In Tolerance Duration") Then Check_Config = False
If Not Range_Check(txtTimeoutDuration(pasTEMPERATURE), 1, 900#, "PAS Temp Timeout Duration") Then Check_Config = False
If Not Range_Check(txtInTolDuration(pasMOISTURE), 1, 900#, "PAS Moisture In Tolerance Duration") Then Check_Config = False
If Not Range_Check(txtTimeoutDuration(pasMOISTURE), 1, 900#, "PAS Moisture Timeout Duration") Then Check_Config = False
    
' Heading Lines
If Len(txtHeading.text) > 60 Then
  txtHeading.BackColor = PALEYELLOW
  lblMessage.Caption = lblMessage.Caption & "SysConfig.Heading Line 1:  60 Characters Max!" & vbCrLf
End If
If Len(txtHeading2.text) > 60 Then
  txtHeading2.BackColor = PALEYELLOW
  lblMessage.Caption = lblMessage.Caption & "SysConfig.Heading Line 2:  60 Characters Max!" & vbCrLf
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

Function Range_Check(tcontrol As Control, slow, shigh As Single, _
slabel As String) As Boolean
' Function Name:    Range_Check
' Author;           Analytical Process Programmer     7/25/96
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

SetErrModule 53, 3
Range_Check = True
If UseLocalErrorHandler Then On Error GoTo localhandler

svalue = CSng(tcontrol.text)

If svalue < slow Or svalue > shigh Then
    Range_Check = False
    tcontrol.BackColor = PALEYELLOW
'    tcontrol.SelStart = 0
'    tcontrol.SelLength = Len(tcontrol.text)
'    tcontrol.SetFocus
    Message = slabel & ": Value out of range!" & " ( " & slow & " - " & shigh & " )"
    lblMessage.Caption = lblMessage.Caption & Message & vbCrLf
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

Sub Update_Config()

' Procedure Name:   Update_Config
' Written By:       Analytical Process Programmer
' Description:
' This procedure updates the user configuration values from the
' configuration data.  This routine does not read or write data
' to a file.
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 53, 5
Dim hilim As Single
    txtFileName.text = Format(SysConfig.Next_File, "000000")    ' File name Seed
    SelectUserName.ListIndex = IIf((SysConfig.AutoLogon > SelectUserName.ListCount - 1), 0, SysConfig.AutoLogon)
'    cboAutoPrint.ListIndex = SysConfig.AutoPrint
'    cboAutoPrint.Refresh
    txtMixRatio.text = Format(SysConfig.Tol_Mix_Ratio, "##0.0")
    txtLoadPressure.text = Format(SysConfig.LoadPressure, "#0.0#")
    txtButaneMassLimit.text = Format(SysConfig.ButaneMassLimit, "#0.0#")
    txtLoadTimeLimit.text = Format(SysConfig.LoadTimeLimit, "#0.0#")
    txtPurgeTotal.text = Format(SysConfig.Tol_Purge_Total, "##0.00")
    txtLoadTotal.text = Format(SysConfig.Tol_Load_Total, "##0.00")
    txtNitTol.text = Format(SysConfig.Tol_Nit_Flow, "#0.0#")
    txtBtnTol.text = Format(SysConfig.Tol_Btn_Flow, "#0.0#")
    txtORVRNitTol.text = Format(SysConfig.Tol_ORVRNit_Flow, "#0.0#")
    txtORVRBtnTol.text = Format(SysConfig.Tol_ORVRBtn_Flow, "#0.0#")
    txtPurgeTol.text = Format(SysConfig.Tol_Pur_Flow, "#0.0#")
    txtLfvTol.text = Format(SysConfig.Tol_Lfv_Flow, "##0.0")
    txtFuelTempTol.text = Format(SysConfig.Tol_FuelTemp, "#0.0")
    txtOvenTempTol.text = Format(SysConfig.Tol_PurgeOvenTemp, "#0.0")
    txtWaterBathTol.text = Format(SysConfig.Tol_WaterBathTemp, "#0.0")
    txtLoLimLoad.text = Format(SysConfig.LoLim_Load_Flow, "0.0")
    txtLoLimPurge.text = Format(SysConfig.LoLim_Purge_Flow, "0.0")
    ' default report interval value
    txtDefaultIntvl.text = Format(SysConfig.Default_Interval, "###0")
    ' load report interval value
    txtLoadIntvl.text = Format(SysConfig.Load_Interval, "###0")
    txtLoadIntvl.ToolTipText = Format(MinDataLogSeconds, "#0.0#") & " to 900 seconds"
    ' purge report interval value
    txtPurgeIntvl.text = Format(SysConfig.Purge_Interval, "###0")
    txtPurgeIntvl.ToolTipText = Format(MinDataLogSeconds, "#0.0#") & " to 900 seconds"
    ' load totalize interval value (must not be greater than Load Report Interval)
    txtLoadTotIntvl.text = Format(SysConfig.LoadTotal_Interval, "###0.0##")
    hilim = 10
    If (hilim > ValueFromText(txtLoadIntvl.text)) Then hilim = ValueFromText(txtLoadIntvl.text)
    txtLoadTotIntvl.ToolTipText = Format(MinDataLogSeconds, "#0.0#") & " to " & Format(hilim, "#0.0#") & " seconds"
    ' purge totalize interval value (must not be greater than Purge Report Interval)
    txtPurgeTotIntvl.text = Format(SysConfig.PurgeTotal_Interval, "###0.0##")
    ' remote status update interval value
    txtRemStatusLogInterval.text = Format(SysConfig.RemStatus_Interval, "###0.0##")
    hilim = 10
    If (hilim > ValueFromText(txtPurgeIntvl.text)) Then hilim = ValueFromText(txtPurgeIntvl.text)
    txtPurgeTotIntvl.ToolTipText = Format(MinDataLogSeconds, "#0.0#") & " to " & Format(hilim, "#0.0#") & " seconds"
        
    txtTempTol.text = Format(SysConfig.Tol_Temp, "#0.0")
    txtMoistureTol.text = Format(SysConfig.Tol_Moisture, "#0.0")
    txtTempTarget.text = Format(SysConfig.Temp_Target, "##0.0")
    txtMoistureTarget.text = Format(SysConfig.Moisture_Target, "##0.0")
    txtPurgeDpHiLimit.text = Format(SysConfig.PurgeDP_HiLimit, "#0.0#")
    txtHeading.text = SysConfig.Heading
    txtHeading2.text = SysConfig.Heading2
    optDbfBackup = IIf(SysConfig.DbFileBackup_Active, 1, 0)
    txtDbfBackupPath.text = SysConfig.DbFileBackup_Path
    optRptBackup = IIf(SysConfig.ReportBackup_Active, 1, 0)
    txtRptBackupPath.text = SysConfig.ReportBackup_Path
    'cboReportName1.ListIndex = SysConfig.ReportFileName1stPart
    cboReportName1.ListIndex = IIf(SysConfig.ReportFileName1stPart > 0, SysConfig.ReportFileName1stPart - 1, 0)
    cboReportName1.Refresh
    cboReportName2.ListIndex = SysConfig.ReportFileName2ndPart
    cboReportName2.Refresh
    cboReportName3.ListIndex = SysConfig.ReportFileName3rdPart
    cboReportName3.Refresh
    txtEventRecs.text = Format(SysConfig.EventRecs, "###0")
    txtJobRecs.text = Format(SysConfig.JobRecs, "###0")
    txtLCMinDelay.text = Format(SysConfig.LCMinDelay, "##0")
    txtLCSetPoint.text = Format(SysConfig.LCSetPoint, "##0.0")
    txtLCTime.text = Format(SysConfig.LCTime, "##0")
    txtPressureDecay.text = Format(SysConfig.PressureDecay, "#0.0")
    txtNitrogenPurgeTime.text = SysConfig.NitrogenPurgeTime
    txtDoorOpenDelay.text = Format(SysConfig.DoorOpenDelay, "##0")
    txtUPSOpenDelay.text = Format(SysConfig.UPSOpenDelay, "##0")
    txtOOTtime.text = Format(SysConfig.OOTtimeDelay, "##0")
    txtLCIntvl.text = Format(SysConfig.LeakCheck_Interval, "###0")
    txtLoadSettleTime.text = Format(SysConfig.LoadSettleTime, "##0.00")
    txtPurgeSettleTime.text = Format(SysConfig.PurgeSettleTime, "##0.00")
    txtOvenBand.text = Format(SysConfig.PurgeOvenBand, "##0.000")
    WaterBathTemperatureControl.ListIndex = SysConfig.WaterBathControl
    WaterBathTemperatureControl.Refresh
    
    ' TEMPERATURE
    If USINGC Then
        lblTempTargetUnits.Caption = "(deg C)"
        lblTempTolUnits.Caption = "(deg C)"
        If USINGHIGHTEMPPAS Then
            txtTempTarget.ToolTipText = "5 to 60 deg C"
            txtTempTol.ToolTipText = "0 to 100 deg C"
        Else
            txtTempTarget.ToolTipText = "5 to 35 deg C"
            txtTempTol.ToolTipText = "0 to 100 deg C"
        End If
        lblOvenBandUnits.Caption = "(deg C)"
        txtOvenBand.ToolTipText = "2 to 5 deg C"
        lblOvenTempUnits.Caption = "(deg C)"
        txtOvenTempTol.ToolTipText = "0 to 100 deg C"
        lblFuelTempUnits.Caption = "(deg C)"
        txtFuelTempTol.ToolTipText = "1 to 10 deg C"
        lblWaterBathUnits.Caption = "(deg C)"
        txtWaterBathTol.ToolTipText = "1 to 10 deg C"
        lblFuelTempUnits.Caption = "(deg C)"
        txtFuelTempTol.ToolTipText = "1 to 10 deg C"
    ElseIf USINGF Then
        lblTempTargetUnits.Caption = "(deg F)"
        lblTempTolUnits.Caption = "(deg F)"
        If USINGHIGHTEMPPAS Then
            txtTempTarget.ToolTipText = "40 to 140 deg F"
            txtTempTol.ToolTipText = "0 to 100 deg F"
        Else
            txtTempTarget.ToolTipText = "40 to 90 deg F"
            txtTempTol.ToolTipText = "0 to 100 deg F"
        End If
        lblOvenBandUnits.Caption = "(deg F)"
        txtOvenBand.ToolTipText = "3.6 to 8 deg F"
        lblOvenTempUnits.Caption = "(deg F)"
        txtOvenTempTol.ToolTipText = "0 to 100 deg F"
        lblFuelTempUnits.Caption = "(deg F)"
        txtFuelTempTol.ToolTipText = "2 to 15 deg F"
        lblWaterBathUnits.Caption = "(deg F)"
        txtWaterBathTol.ToolTipText = "2 to 15 deg F"
        lblFuelTempUnits.Caption = "(deg F)"
        txtFuelTempTol.ToolTipText = "2 to 15 deg F"
    End If
    
If USINGC Then
        If USINGHIGHTEMPPAS Then
            txtTempTarget.ToolTipText = "5 to 60 deg C"
            txtTempTol.ToolTipText = "0 to 100 deg C"
        Else
            txtTempTarget.ToolTipText = "5 to 35 deg C"
            txtTempTol.ToolTipText = "0 to 100 deg C"
        End If
ElseIf USINGF Then
        If USINGHIGHTEMPPAS Then
            txtTempTarget.ToolTipText = "40 to 140 deg F"
            txtTempTol.ToolTipText = "0 to 100 deg F"
        Else
            txtTempTarget.ToolTipText = "40 to 90 deg F"
            txtTempTol.ToolTipText = "0 to 100 deg F"
        End If
End If
    
    
    ' MOISTURE
    If USINGMoist_RH Then
        lblMoistTargetUnits.Caption = "(% rH)"
        txtMoistureTarget.ToolTipText = "0 to 100% rH"
        lblMoistTolUnits.Caption = "(% rH)"
        txtMoistureTol.ToolTipText = "0 to 100% rH"
    ElseIf USINGMoist_Grains Then
        lblMoistTargetUnits.Caption = "(grains/lb)"
        txtMoistureTarget.ToolTipText = "0 to 200 grains/lb"
        lblMoistTolUnits.Caption = "(grains/lb)"
        txtMoistureTol.ToolTipText = "0 to 100 grains/lb"
    End If
            
    ' Temp & Humidity Logging
    txtLogTempRhInterval.text = Format(SysConfig.TempRhLogInterval, "##0")
    optLogTempRhVerbose.Value = IIf(SysConfig.TempRhLogVerbose, 1, 0)
            
    lblLCSetPoint.Caption = "Leak Check Pressure Set Point:   (psig)"
    txtLCSetPoint.ToolTipText = "0.1 to 15 psig"
    
    ' Leakcheck Fail Response
    SelectLkChkFailResp.ListIndex = SysConfig.LeakCheckFailResponse
    
    ' Out-Of-Tolerance Response
    ResponseOOT(ootBtnFlow).ListIndex = SysConfig.BtnFlowResp - 1
    ResponseOOT(ootNitFlow).ListIndex = SysConfig.NitFlowResp - 1
    ResponseOOT(ootStorageLevel).ListIndex = SysConfig.StorageLevelResp - 1
    ResponseOOT(ootFuelLevel).ListIndex = SysConfig.FuelLevelResp - 1
    ResponseOOT(ootFuelTemp).ListIndex = SysConfig.FuelTempResp - 1
    ResponseOOT(ootPurFlow).ListIndex = SysConfig.PurFlowResp - 1
    ResponseOOT(ootAirMoist).ListIndex = SysConfig.AirMoistResp - 1
    ResponseOOT(ootAirTemp).ListIndex = SysConfig.AirTempResp - 1
    ResponseOOT(ootCanVent).ListIndex = SysConfig.CanVentResp - 1
    ResponseOOT(ootLoadRate).ListIndex = SysConfig.LoadRateResp - 1
    ResponseOOT(ootPurgeDp).ListIndex = SysConfig.PurgeDpResp - 1
    ResponseOOT(ootPurgeOvenTemp).ListIndex = SysConfig.PurgeOvenResp - 1
    ResponseOOT(ootWaterBathTemp).ListIndex = SysConfig.WaterBathResp - 1
        
    chkPosPressPurge.Value = IIf(SysConfig.PosPressPurge, cYES, cNO)
    chkDryAirPurge.Value = IIf(SysConfig.DryAirPurge, cYES, cNO)
    
    If USINGUPS = 2 Then
       txtUPSOpenDelay.ToolTipText = "Max delay from 1 to 10 minutes"
    End If
    txtCanventOvr.text = Format(SysConfig.CanVent_Delay_Max, "##0")
    
    txtOnDutyMult(pasTEMPERATURE).text = Format(PID_INFO(pasTEMPERATURE).OnDutyMult, "0.00")
    txtOffDutyMult(pasTEMPERATURE).text = Format(PID_INFO(pasTEMPERATURE).OffDutyMult, "0.00")
    txtInTolDuration(pasTEMPERATURE).text = Format(PAS_INFO(pasTEMPERATURE).DurationTarget, "###0")
    txtTimeoutDuration(pasTEMPERATURE).text = Format(PAS_INFO(pasTEMPERATURE).TimeOutTarget, "###0")
    txtPgain(pasMOISTURE).text = Format(PID_INFO(pasMOISTURE).Pgain, "##0.00")
    txtIgain(pasMOISTURE).text = Format(PID_INFO(pasMOISTURE).Igain, "##0.00")
    txtInTolDuration(pasMOISTURE).text = Format(PAS_INFO(pasMOISTURE).DurationTarget, "###0")
    txtTimeoutDuration(pasMOISTURE).text = Format(PAS_INFO(pasMOISTURE).TimeOutTarget, "###0")
    
    
    ' Report Configuration
    With SysConfig.RptConfig
        ' value
        chkCsvEotReporting.Value = IIf(.CsvEotReporting, cYES, cNO)
        chkCsvEotSummary.Value = IIf(.CsvEotSummary, cYES, cNO)
        chkCsvEotDetail.Value = IIf(.CsvEotDetail, cYES, cNO)
        chkCsvGenReporting.Value = IIf(.CsvGenReporting, cYES, cNO)
        chkCsvGenSummary.Value = IIf(.CsvGenSummary, cYES, cNO)
        chkCsvGenDetail.Value = IIf(.CsvGenDetail, cYES, cNO)
        chkTextEotReporting.Value = IIf(.TextEotReporting, cYES, cNO)
        chkTextEotSummary.Value = IIf(.TextEotSummary, cYES, cNO)
        chkTextEotSummaryAutoPrint.Value = IIf(.TextEotSummary_AutoPrint, cYES, cNO)
        chkTextEotDetail.Value = IIf(.TextEotDetail, cYES, cNO)
        chkTextGenReporting.Value = IIf(.TextGenReporting, cYES, cNO)
        chkTextGenSummary.Value = IIf(.TextGenSummary, cYES, cNO)
        chkTextGenDetail.Value = IIf(.TextGenDetail, cYES, cNO)
        chkXlsEotReporting.Value = IIf(.XlsEotReporting, cYES, cNO)
        chkXlsEotSummary.Value = IIf(.XlsEotSummary, cYES, cNO)
        chkXlsEotDetail.Value = IIf(.XlsEotDetail, cYES, cNO)
        chkXlsGenReporting.Value = IIf(.XlsGenReporting, cYES, cNO)
        chkXlsGenSummary.Value = IIf(.XlsGenSummary, cYES, cNO)
        chkXlsGenDetail.Value = IIf(.XlsGenDetail, cYES, cNO)
    End With
    Update_RptConfigAppearance
  
    ' AutoDrainFill
    If (LiveFuelStn > 0) Then Update_ADF LiveFuelStn

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

Sub Update_RptConfigAppearance()
' Procedure Name:   Update_RptConfigAppearance
' Written By:       Brunrose
' Description:
' This procedure updates the appearance of
' the user report configuration selection items
' on the Config screen
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 53, 52

    With SysConfig.RptConfig
        ' enabled
        chkTextEotReporting.Enabled = True
        chkTextEotSummary.Enabled = IIf((chkTextEotReporting.Value = cYES), True, False)
        chkTextEotSummaryAutoPrint.Enabled = IIf(((chkTextEotReporting.Value = cYES) And (chkTextEotSummary.Value = cYES)), True, False)
        chkTextEotDetail.Enabled = IIf((chkTextEotReporting.Value = cYES), True, False)
        chkXlsEotReporting.Enabled = True
        chkXlsEotSummary.Enabled = IIf((chkXlsEotReporting.Value = cYES), True, False)
        chkXlsEotDetail.Enabled = IIf((chkXlsEotReporting.Value = cYES), True, False)
        chkCsvEotReporting.Enabled = True
        chkCsvEotSummary.Enabled = IIf((chkCsvEotReporting.Value = cYES), True, False)
        chkCsvEotDetail.Enabled = IIf((chkCsvEotReporting.Value = cYES), True, False)
        chkTextGenReporting.Enabled = True
        chkTextGenSummary.Enabled = IIf((chkTextGenReporting.Value = cYES), True, False)
        chkTextGenDetail.Enabled = IIf((chkTextGenReporting.Value = cYES), True, False)
        chkXlsGenReporting.Enabled = True
        chkXlsGenSummary.Enabled = IIf((chkXlsGenReporting.Value = cYES), True, False)
        chkXlsGenDetail.Enabled = IIf((chkXlsGenReporting.Value = cYES), True, False)
        chkCsvGenReporting.Enabled = True
        chkCsvGenSummary.Enabled = IIf((chkCsvGenReporting.Value = cYES), True, False)
        chkCsvGenDetail.Enabled = IIf((chkCsvGenReporting.Value = cYES), True, False)
        ' visible
        chkTextEotReporting.Visible = True
        chkTextEotSummary.Visible = True
        chkTextEotSummaryAutoPrint.Visible = True
        chkTextEotDetail.Visible = True
        chkTextGenReporting.Visible = True
        chkTextGenSummary.Visible = True
        chkTextGenDetail.Visible = True
        chkXlsGenReporting.Visible = True
        chkXlsGenSummary.Visible = True
        chkXlsGenDetail.Visible = True
        chkCsvGenReporting.Visible = True
        chkCsvGenSummary.Visible = True
        chkCsvGenDetail.Visible = True
    End With
  
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

Sub Refresh_Config()

' Procedure Name:   Refresh_config
' Created By:       Analytical Process Programmer
' Description:
' This procedure copies the user configuration data from the config screen
' to the configuration data variables.
'
' First Update memory variables from form variables
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 53, 4

    SysConfig.Next_File = txtFileName.text
    SysConfig.AutoLogon = SelectUserName.ListIndex
    SysConfig.AutoLogonUser = SelectUserName.List(SelectUserName.ListIndex)
'    SysConfig.AutoPrint = cboAutoPrint.ListIndex
    SysConfig.Tol_Mix_Ratio = ValueFromText(txtMixRatio.text)
    SysConfig.LoadPressure = ValueFromText(txtLoadPressure.text)
    SysConfig.ButaneMassLimit = ValueFromText(txtButaneMassLimit.text)
    SysConfig.LoadTimeLimit = ValueFromText(txtLoadTimeLimit.text)
    SysConfig.LoLim_Load_Flow = ValueFromText(txtLoLimLoad.text)              ' Low Limit for Tolerance Checking in %
    SysConfig.LoLim_Purge_Flow = ValueFromText(txtLoLimPurge.text)            ' Low Limit for Tolerance Checking in %
    SysConfig.Tol_Load_Total = ValueFromText(txtLoadTotal.text)
    SysConfig.Tol_Purge_Total = ValueFromText(txtPurgeTotal.text)
    SysConfig.Tol_Nit_Flow = ValueFromText(txtNitTol.text)
    SysConfig.Tol_Btn_Flow = ValueFromText(txtBtnTol.text)
    SysConfig.Tol_ORVRNit_Flow = ValueFromText(txtORVRNitTol.text)
    SysConfig.Tol_ORVRBtn_Flow = ValueFromText(txtORVRBtnTol.text)
    SysConfig.Tol_Pur_Flow = ValueFromText(txtPurgeTol.text)
    SysConfig.Tol_Lfv_Flow = ValueFromText(txtLfvTol.text)
    SysConfig.Tol_FuelTemp = ValueFromText(txtFuelTempTol.text)
    SysConfig.Tol_PurgeOvenTemp = ValueFromText(txtOvenTempTol.text)
    SysConfig.Tol_WaterBathTemp = ValueFromText(txtWaterBathTol.text)
    SysConfig.PurgeDP_HiLimit = ValueFromText(txtPurgeDpHiLimit.text)
    SysConfig.Default_Interval = CInt(txtDefaultIntvl.text)
    SysConfig.Load_Interval = CInt(txtLoadIntvl.text)
    SysConfig.Purge_Interval = CInt(txtPurgeIntvl.text)
    SysConfig.LoadTotal_Interval = CSng(txtLoadTotIntvl.text)
    SysConfig.PurgeTotal_Interval = CSng(txtPurgeTotIntvl.text)
    SysConfig.RemStatus_Interval = CSng(txtRemStatusLogInterval.text)
    SysConfig.Tol_Temp = ValueFromText(txtTempTol.text)
    SysConfig.Tol_Moisture = ValueFromText(txtMoistureTol.text)
    SysConfig.Temp_Target = ValueFromText(txtTempTarget.text)
    SysConfig.Moisture_Target = ValueFromText(txtMoistureTarget.text)
    SysConfig.Heading = txtHeading.text
    SysConfig.Heading2 = txtHeading2.text
    SysConfig.DbFileBackup_Active = IIf((optDbfBackup = 1), True, False)
    SysConfig.DbFileBackup_Path = txtDbfBackupPath.text
    SysConfig.ReportBackup_Active = IIf((optRptBackup = 1), True, False)
    SysConfig.ReportBackup_Path = txtRptBackupPath.text
    SysConfig.ReportFileName1stPart = cboReportName1.ListIndex + 1
    SysConfig.ReportFileName2ndPart = cboReportName2.ListIndex
    SysConfig.ReportFileName3rdPart = cboReportName3.ListIndex
    SysConfig.EventRecs = CInt(txtEventRecs.text)
    SysConfig.JobRecs = CInt(txtJobRecs.text)
    SysConfig.LCMinDelay = CInt(txtLCMinDelay.text)
    SysConfig.LCSetPoint = ValueFromText(txtLCSetPoint.text)
    SysConfig.LCTime = CInt(txtLCTime.text)
    SysConfig.PressureDecay = ValueFromText(txtPressureDecay.text)
    SysConfig.NitrogenPurgeTime = CInt(txtNitrogenPurgeTime.text)
    SysConfig.DoorOpenDelay = CInt(txtDoorOpenDelay.text)
    SysConfig.UPSOpenDelay = CInt(txtUPSOpenDelay.text)
    SysConfig.OOTtimeDelay = CInt(txtOOTtime.text)
    SysConfig.CanVent_Delay_Max = CInt(txtCanventOvr.text)
    SysConfig.LeakCheck_Interval = CInt(txtLCIntvl.text)
    SysConfig.LoadSettleTime = ValueFromText(txtLoadSettleTime.text)
    SysConfig.PurgeSettleTime = ValueFromText(txtPurgeSettleTime.text)
    SysConfig.PurgeOvenBand = ValueFromText(txtOvenBand.text)
    
   ' Leakcheck Failure Response
    If USINGCONTAFTERLCFAIL Then
        SysConfig.LeakCheckFailResponse = SelectLkChkFailResp.ListIndex
    Else
        SysConfig.LeakCheckFailResponse = 0
    End If
    
    If USINGPRESSUREPURGE Then
        SysConfig.PosPressPurge = IIf((chkPosPressPurge.Value = cYES), True, False)
    Else
        SysConfig.PosPressPurge = False
    End If
    
    If USINGDRYPURGEAIR Then
        SysConfig.DryAirPurge = IIf((chkDryAirPurge.Value = cYES), True, False)
    Else
        SysConfig.DryAirPurge = False
    End If
    
    If USINGWATERBATH Then
        SysConfig.WaterBathControl = WaterBathTemperatureControl.ListIndex
    Else
        SysConfig.WaterBathControl = wbDirect
    End If
    
   ' Out-Of-Tolerance Response
   SysConfig.BtnFlowResp = ResponseOOT(ootBtnFlow).ListIndex + 1
   SysConfig.NitFlowResp = ResponseOOT(ootNitFlow).ListIndex + 1
   SysConfig.StorageLevelResp = ResponseOOT(ootStorageLevel).ListIndex + 1
   SysConfig.FuelLevelResp = ResponseOOT(ootFuelLevel).ListIndex + 1
   SysConfig.FuelTempResp = ResponseOOT(ootFuelTemp).ListIndex + 1
   SysConfig.PurFlowResp = ResponseOOT(ootPurFlow).ListIndex + 1
   SysConfig.AirMoistResp = ResponseOOT(ootAirMoist).ListIndex + 1
   SysConfig.AirTempResp = ResponseOOT(ootAirTemp).ListIndex + 1
   SysConfig.CanVentResp = ResponseOOT(ootCanVent).ListIndex + 1
   SysConfig.LoadRateResp = ResponseOOT(ootLoadRate).ListIndex + 1
   SysConfig.PurgeDpResp = ResponseOOT(ootPurgeDp).ListIndex + 1
   SysConfig.PurgeOvenResp = ResponseOOT(ootPurgeOvenTemp).ListIndex + 1
   SysConfig.WaterBathResp = ResponseOOT(ootWaterBathTemp).ListIndex + 1
    
    ' Temp & Humidity Logging
    SysConfig.TempRhLogInterval = ValueFromText(txtLogTempRhInterval.text)
    SysConfig.TempRhLogVerbose = IIf((optLogTempRhVerbose.Value = 1), True, False)
    
    ' PAS Local Control
    If USINGPASLOCALCONTROL Then
        ' temperature controller
        PID_INFO(pasTEMPERATURE).OnDutyMult = ValueFromText(txtOnDutyMult(pasTEMPERATURE).text)
        PID_INFO(pasTEMPERATURE).OffDutyMult = ValueFromText(txtOffDutyMult(pasTEMPERATURE).text)
        PAS_INFO(pasTEMPERATURE).DurationTarget = CDbl(txtInTolDuration(pasTEMPERATURE).text)
        PAS_INFO(pasTEMPERATURE).TimeOutTarget = CDbl(txtTimeoutDuration(pasTEMPERATURE).text)
        ' moisture controller
        PID_INFO(pasMOISTURE).Pgain = ValueFromText(txtPgain(pasMOISTURE).text)
        PID_INFO(pasMOISTURE).Igain = ValueFromText(txtIgain(pasMOISTURE).text)
        PAS_INFO(pasMOISTURE).DurationTarget = CDbl(txtInTolDuration(pasMOISTURE).text)
        PAS_INFO(pasMOISTURE).TimeOutTarget = CDbl(txtTimeoutDuration(pasMOISTURE).text)
    End If
    
    ' Report Configuration
    With SysConfig.RptConfig
        .CsvEotReporting = IIf((chkCsvEotReporting.Value = cYES), True, False)
        .CsvEotSummary = IIf((chkCsvEotSummary.Value = cYES), True, False)
        .CsvEotDetail = IIf((chkCsvEotDetail.Value = cYES), True, False)
        .CsvGenReporting = IIf((chkCsvGenReporting.Value = cYES), True, False)
        .CsvGenSummary = IIf((chkCsvGenSummary.Value = cYES), True, False)
        .CsvGenDetail = IIf((chkCsvGenDetail.Value = cYES), True, False)
        .TextEotReporting = IIf((chkTextEotReporting.Value = cYES), True, False)
        .TextEotSummary = IIf((chkTextEotSummary.Value = cYES), True, False)
        .TextEotSummary_AutoPrint = IIf((chkTextEotSummaryAutoPrint.Value = cYES), True, False)
        .TextEotDetail = IIf((chkTextEotDetail.Value = cYES), True, False)
        .TextGenReporting = IIf((chkTextGenReporting.Value = cYES), True, False)
        .TextGenSummary = IIf((chkTextGenSummary.Value = cYES), True, False)
        .TextGenDetail = IIf((chkTextGenDetail.Value = cYES), True, False)
        .XlsEotReporting = IIf((chkXlsEotReporting.Value = cYES), True, False)
        .XlsEotSummary = IIf((chkXlsEotSummary.Value = cYES), True, False)
        .XlsEotDetail = IIf((chkXlsEotDetail.Value = cYES), True, False)
        .XlsGenReporting = IIf((chkXlsGenReporting.Value = cYES), True, False)
        .XlsGenSummary = IIf((chkXlsGenSummary.Value = cYES), True, False)
        .XlsGenDetail = IIf((chkXlsGenDetail.Value = cYES), True, False)
    End With

    ' LiveFuel AutoDrainFill
    If systemhasLIVEFUEL And LiveFuelStn >= 0 And LiveFuelStn <= NR_STN Then
        StationCfg_ADF(LiveFuelStn, 1).DrainDelay = CInt(txtLiveFuelChgDrainDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).DrainTimeout = CInt(txtLiveFuelChgDrainTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).DrainShutOff = CSng(txtLiveFuelChgDrainShutoff.text)
        StationCfg_ADF(LiveFuelStn, 1).FillDelay = CInt(txtLiveFuelChgFillDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).FillTimeout = CInt(txtLiveFuelChgFillTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).FillShutOff = CSng(txtLiveFuelChgFillShutoff.text)
        StationCfg_ADF(LiveFuelStn, 1).PurgeDrainDelay = CInt(txtLiveFuelChgPurgeDrainDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).PurgeFillDelay = CInt(txtLiveFuelChgPurgeFillDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).PurgeTimeout = CInt(txtLiveFuelChgPurgeTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).HeaterTimeout = CInt(txtLiveFuelChgHeaterTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).VaporGenTankVol = CSng(txtVaporGenTankVol.text)
        StationCfg_ADF(LiveFuelStn, 1).VaporGenLevelTol = CSng(txtVaporGenLeakRate.text)
        StationCfg_ADF(LiveFuelStn, 1).FuelStorageTankVol = CSng(txtFuelStorageTankVol.text)
        StationCfg_ADF(LiveFuelStn, 1).FuelStorageLevelTol = CSng(txtFuelStorageLeakRate.text)
        StationCfg_ADF(LiveFuelStn, 1).FstDrainDelay = CInt(txtFuelStorageDrainDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).FstDrainTimeout = CInt(txtFuelStorageDrainTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).FstDrainShutOff = CSng(txtFuelStorageDrainShutoff.text)
        StationCfg_ADF(LiveFuelStn, 1).FstFillDelay = CInt(txtFuelStorageFillDelay.text)
        StationCfg_ADF(LiveFuelStn, 1).FstFillTimeout = CInt(txtFuelStorageFillTimeout.text)
        StationCfg_ADF(LiveFuelStn, 1).FstFillShutOff = CSng(txtFuelStorageFillShutoff.text)
        PID_INFO(LiveFuelStn + 10).Pgain = CSng(txtLoadRate_Pgain.text)
        PID_INFO(LiveFuelStn + 10).Igain = CSng(txtLoadRate_Igain.text)
    End If
  
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

Private Sub cfgtabs_Click(PreviousTab As Integer)
    Select Case cfgtabs.Tab
        Case Tab_OotResponse
            ' align OOT Response boxes
            frmOotResp(ootBtnFlow).Left = OotCol1Left
            frmOotResp(ootNitFlow).Left = OotCol1Left
            frmOotResp(ootFuelTemp).Left = OotCol1Left
            frmOotResp(ootLoadRate).Left = OotCol1Left
            frmOotResp(ootStorageLevel).Left = OotCol2Left
            frmOotResp(ootFuelLevel).Left = OotCol2Left
            frmOotResp(ootCanVent).Left = OotCol2Left
            frmOotResp(ootPurFlow).Left = OotCol3Left
            frmOotResp(ootAirMoist).Left = OotCol3Left
            frmOotResp(ootAirTemp).Left = OotCol3Left
            frmOotResp(ootPurgeDp).Left = OotCol3Left
        Case Else
            ' nothing to do
    End Select
End Sub

Private Sub chkCsvEotReporting_Click()
    If (chkCsvEotReporting.Value = cNO) Then
        chkCsvEotSummary.Value = cNO
        chkCsvEotDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub chkCsvGenReporting_Click()
    If (chkCsvGenReporting.Value = cNO) Then
        chkCsvGenSummary.Value = cNO
        chkCsvGenDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub chkTextEotReporting_Click()
    If (chkTextEotReporting.Value = cNO) Then
        chkTextEotSummary.Value = cNO
        chkTextEotDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub chkTextEotSummary_Click()
    Update_RptConfigAppearance
End Sub

Private Sub chkTextGenReporting_Click()
    If (chkTextGenReporting.Value = cNO) Then
        chkTextGenSummary.Value = cNO
        chkTextGenDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub chkXlsEotReporting_Click()
    If (chkXlsEotReporting.Value = cNO) Then
        chkXlsEotSummary.Value = cNO
        chkXlsEotDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub chkXlsGenReporting_Click()
    If (chkXlsGenReporting.Value = cNO) Then
        chkXlsGenSummary.Value = cNO
        chkXlsGenDetail.Value = cNO
    End If
    Update_RptConfigAppearance
End Sub

Private Sub cmdPrint_Click()
    Print_Config
'    Delay_Box "Configuration screens Released to the Printer", MSGDELAY, msgSHOW
    lblMessage.Font.Size = 9.5
    lblMessage.ForeColor = DKPURPLE
    lblMessage.Caption = vbCrLf & "Configuration screens sent to" & vbCrLf & PRINTERNAME
End Sub

Private Sub cmdReturn_Click()
    Unload Me
    Set frmConfig = Nothing   'testing
End Sub

Private Sub cmdSave_Click()
    If CheckPass("O", True) Then
        lblMessage.Caption = vbCrLf
        If Check_Config Then
            Refresh_Config          ' Copy data to configuration variables
            Save_Config             ' Save configuration data to disk
            If USINGREMCANLOAD Then
                ' open master canister / recipe database
                Set dbDbase = OpenDatabase(FILEPATH_rcp & DATARCP)
                ' open remote database
                OpenConnToRemoteDb
                ' update Remote Configuration Information
                UpdateRemoteConfiguration
                ' close remote database
                CloseConnToRemoteDb
            End If
            Save_AdfConfig
            Save_LocalPAS
            Save_Controllers        ' Save Controllers Config Data
            Update_Config           ' Update display of configuration variables
    '        Delay_Box "Configuration File Saved", MSGDELAY, msgSHOW
            lblMessage.ForeColor = Message_ForeColor
            lblMessage.Caption = lblMessage.Caption & vbCrLf
            lblMessage.Caption = lblMessage.Caption & "New Configuration Values Saved" & vbCrLf
        Else
            lblMessage.ForeColor = Alarm_ForeColor
            lblMessage.Caption = lblMessage.Caption & vbCrLf
            lblMessage.Caption = lblMessage.Caption & "Configuration Values Not Saved" & vbCrLf
            lblMessage.Caption = lblMessage.Caption & "Try again after correcting the errors." & vbCrLf
            Beep
            Beep
            Beep
        End If
    End If
End Sub

Private Sub cmdStnDn_Click()
Dim iStn As Integer
Dim CurStn As Integer
Dim LoopCount As Integer
    LoopCount = 0
    CurStn = LiveFuelStn
    iStn = CurStn
    Do While CurStn = LiveFuelStn And LoopCount <= NR_STN
        iStn = iStn - 1
        If iStn < 1 Then iStn = NR_STN
        If STN_INFO(iStn).ADF_TANKTYPE <> 0 Then LiveFuelStn = iStn
        LoopCount = LoopCount + 1
    Loop
    ' display options
    Disp_ADF LiveFuelStn
    ' update text
    Update_ADF LiveFuelStn
End Sub

Private Sub cmdStnUp_Click()
Dim iStn As Integer
Dim CurStn As Integer
Dim LoopCount As Integer
    LoopCount = 0
    CurStn = LiveFuelStn
    iStn = CurStn
    Do While CurStn = LiveFuelStn And LoopCount <= NR_STN
        iStn = iStn + 1
        If iStn > NR_STN Then iStn = 1
        If STN_INFO(iStn).ADF_TANKTYPE <> 0 Then LiveFuelStn = iStn
        LoopCount = LoopCount + 1
    Loop
    ' display options
    Disp_ADF LiveFuelStn
    ' update text
    Update_ADF LiveFuelStn
End Sub

Private Sub Disp_ADF(ByVal iStn As Integer)
Dim heaterflag As Boolean
Dim shutofflag As Boolean
Dim storageflag As Boolean
Dim stnflag As Boolean
Dim tunerflag As Boolean
    ' station prev/next controls
    txtDispStn.text = Format(iStn, "#0")
    stnflag = IIf(NumberOfLiveFuelStations > 1, True, False)
    cmdStnDn.Visible = stnflag
    cmdStnUp.Visible = stnflag
    ' drain/fill delays and timeouts
    txtLiveFuelChgDrainDelay.Enabled = True
    txtLiveFuelChgDrainDelay.Visible = True
    txtLiveFuelChgDrainTimeout.Enabled = True
    txtLiveFuelChgDrainTimeout.Visible = True
    txtLiveFuelChgDrainShutoff.Enabled = True
    txtLiveFuelChgDrainShutoff.Visible = True
    txtLiveFuelChgFillDelay.Enabled = True
    txtLiveFuelChgFillDelay.Visible = True
    txtLiveFuelChgFillTimeout.Enabled = True
    txtLiveFuelChgFillTimeout.Visible = True
    txtLiveFuelChgFillShutoff.Enabled = True
    txtLiveFuelChgFillShutoff.Visible = True
    lblADF_PurgeDrainDelay.Visible = True
    lblADF_PurgeDrainDelay2.Visible = True
    txtLiveFuelChgPurgeDrainDelay.Enabled = True
    txtLiveFuelChgPurgeDrainDelay.Visible = True
    lblADF_PurgeFillDelay.Visible = True
    lblADF_PurgeFillDelay2.Visible = True
    txtLiveFuelChgPurgeFillDelay.Enabled = True
    txtLiveFuelChgPurgeFillDelay.Visible = True
    lblADF_PurgeTimeout.Visible = True
    lblADF_PurgeTimeout2.Visible = True
    txtLiveFuelChgPurgeTimeout.Enabled = True
    txtLiveFuelChgPurgeTimeout.Visible = True
    ' fuel temp & n2 purge options
    Select Case STN_INFO(iStn).ADF_TANKTYPE
        Case 1
            ' Pump, Drain, Fill, No N2Purge, No Heater, Level Switches
            heaterflag = False
            shutofflag = False
            storageflag = False
        Case 12
            ' Pump, Drain, Fill, Vapor, Bypass, N2Purge and Heater, Level Switches
            heaterflag = True
            shutofflag = False
            storageflag = False
        Case 20
            ' Pump, Drain, Fill, Vapor, Level Xmtr, Heater
            heaterflag = True
            shutofflag = True
            storageflag = False
        Case 22
            ' Pump, Drain, Fill, Vapor, Level Xmtr, Level Switches, Storage Tank
            heaterflag = False
            shutofflag = True
            storageflag = True
        Case Else
            ' No Auto Drain/Fill
            heaterflag = False
            shutofflag = False
            storageflag = False
    End Select
    ' heater related options
    txtFuelTempTol.Visible = heaterflag
    lblFuelTempTol.Visible = heaterflag
    lblFuelTempUnits.Visible = heaterflag
    lblADF_HeaterTimeout.Visible = heaterflag
    lblADF_HeaterTimeout2.Visible = heaterflag
    txtLiveFuelChgHeaterTimeout.Visible = heaterflag
    txtLiveFuelChgHeaterTimeout.Enabled = heaterflag
    ' level shutoff options
    lblADF_Shutoff.Visible = shutofflag
    lblADF_Shutoff2.Visible = shutofflag
    txtLiveFuelChgDrainShutoff.Enabled = shutofflag
    txtLiveFuelChgDrainShutoff.Visible = shutofflag
    txtLiveFuelChgFillShutoff.Enabled = shutofflag
    txtLiveFuelChgFillShutoff.Visible = shutofflag
    ' loadrate PID options
    tunerflag = IIf(CheckPass("C", False), True, False)
    lblLoadRate_Pgain.Visible = tunerflag
    txtLoadRate_Pgain.Enabled = tunerflag
    txtLoadRate_Pgain.Visible = tunerflag
    lblLoadRate_Igain.Visible = tunerflag
    txtLoadRate_Igain.Enabled = tunerflag
    txtLoadRate_Igain.Visible = tunerflag
    ' tank options
    txtVaporGenTankVol.Visible = True
    lblVaporGenTankVol.Visible = True
    lblVaporGenTankVol2.Visible = True
    lblVaporGenTankVol2.Caption = IIf(SysSysDef.USINGLVol_SI, "(liters)", "(gallons)")
    txtFuelStorageTankVol.Visible = storageflag
    lblFuelStorageTankVol.Visible = storageflag
    lblFuelStorageTankVol2.Visible = storageflag
    lblFuelStorageTankVol2.Caption = IIf(SysSysDef.USINGLVol_SI, "(liters)", "(gallons)")
    txtVaporGenLeakRate.Visible = True
    lblVaporGenLeakRate.Visible = True
    lblVaporGenLeakRateUnits.Visible = True
    txtFuelStorageLeakRate.Visible = storageflag
    lblFuelStorageLeakRate.Visible = storageflag
    lblFuelStorageLeakRateUnits.Visible = storageflag
    ' storage tank options
    lblStorageTank.Visible = storageflag
    lblFST_Drain.Visible = storageflag
    lblFST_Fill.Visible = storageflag
    lblADF_Shutoff2.Visible = storageflag
    lblFST_Delay.Visible = storageflag
    lblFST_Delay2.Visible = storageflag
    lblFST_Timeout.Visible = storageflag
    lblFST_Timeout2.Visible = storageflag
    lblFST_Shutoff.Visible = storageflag
    lblFST_Shutoff2.Visible = storageflag
    txtFuelStorageDrainDelay.Enabled = storageflag
    txtFuelStorageDrainDelay.Visible = storageflag
    txtFuelStorageFillDelay.Enabled = storageflag
    txtFuelStorageFillDelay.Visible = storageflag
    txtFuelStorageDrainTimeout.Enabled = storageflag
    txtFuelStorageDrainTimeout.Visible = storageflag
    txtFuelStorageFillTimeout.Enabled = storageflag
    txtFuelStorageFillTimeout.Visible = storageflag
    txtFuelStorageDrainShutoff.Enabled = storageflag
    txtFuelStorageDrainShutoff.Visible = storageflag
    txtFuelStorageFillShutoff.Enabled = storageflag
    txtFuelStorageFillShutoff.Visible = storageflag

End Sub

Sub Update_ADF(ByVal iStn As Integer)

' Procedure Name:   Update_ADF
' Written By:       Brunrose
' Description:
' This procedure updates the screen configuration values from the
' AutoDrainFill config data.  This routine does not read or write data
' to a file.
'

SetErrModule 53, 885
If UseLocalErrorHandler Then On Error GoTo localhandler
    
    txtLiveFuelChgDrainDelay.text = Format(StationCfg_ADF(iStn, 1).DrainDelay, "###0")
    txtLiveFuelChgDrainTimeout.text = Format(StationCfg_ADF(iStn, 1).DrainTimeout, "###0")
    txtLiveFuelChgDrainShutoff.text = Format(StationCfg_ADF(iStn, 1).DrainShutOff, "###0")
    txtLiveFuelChgFillDelay.text = Format(StationCfg_ADF(iStn, 1).FillDelay, "###0")
    txtLiveFuelChgFillTimeout.text = Format(StationCfg_ADF(iStn, 1).FillTimeout, "###0")
    txtLiveFuelChgFillShutoff.text = Format(StationCfg_ADF(iStn, 1).FillShutOff, "###0")
    txtLiveFuelChgPurgeDrainDelay.text = Format(StationCfg_ADF(iStn, 1).PurgeDrainDelay, "###0")
    txtLiveFuelChgPurgeFillDelay.text = Format(StationCfg_ADF(iStn, 1).PurgeFillDelay, "###0")
    txtLiveFuelChgPurgeTimeout.text = Format(StationCfg_ADF(iStn, 1).PurgeTimeout, "###0")
    txtLiveFuelChgHeaterTimeout.text = Format(StationCfg_ADF(iStn, 1).HeaterTimeout, "###0")
    txtLoadRate_Pgain.text = Format(PID_INFO(iStn + 10).Pgain, "##0.00")
    txtLoadRate_Igain.text = Format(PID_INFO(iStn + 10).Igain, "##0.00")
    txtVaporGenTankVol.text = Format(StationCfg_ADF(iStn, 1).VaporGenTankVol, "###0.0#")
    txtVaporGenLeakRate.text = Format(StationCfg_ADF(iStn, 1).VaporGenLevelTol, "###0.0##")
    txtFuelStorageTankVol.text = Format(StationCfg_ADF(iStn, 1).FuelStorageTankVol, "###0.0#")
    txtFuelStorageLeakRate.text = Format(StationCfg_ADF(iStn, 1).FuelStorageLevelTol, "###0.0##")
    txtFuelStorageDrainDelay.text = Format(StationCfg_ADF(iStn, 1).FstDrainDelay, "###0")
    txtFuelStorageDrainTimeout.text = Format(StationCfg_ADF(iStn, 1).FstDrainTimeout, "###0")
    txtFuelStorageDrainShutoff.text = Format(StationCfg_ADF(iStn, 1).FstDrainShutOff, "###0")
    txtFuelStorageFillDelay.text = Format(StationCfg_ADF(iStn, 1).FstFillDelay, "###0")
    txtFuelStorageFillTimeout.text = Format(StationCfg_ADF(iStn, 1).FstFillTimeout, "###0")
    txtFuelStorageFillShutoff.text = Format(StationCfg_ADF(iStn, 1).FstFillShutOff, "###0")

    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
        Set frmConfig = Nothing
    End If
End Sub

Private Sub CfgAutoLogonInit()
' Initialize Password Access
Dim pUser As Passkey
Dim dbDbase As Database
Dim rsRecord As Recordset
Dim rsCriterion As String
Dim sPath, sUserName As String
Dim Idx As Integer

    ' Check password list
    Set dbDbase = OpenDatabase(FILEPATH_sysdbf & DATAUSER)
    rsCriterion = "SELECT * FROM [Password] ORDER BY [Password].[UserName] ASC"
    Set rsRecord = dbDbase.OpenRecordset(rsCriterion, dbOpenDynaset)
    If Not rsRecord.BOF Then  ' See if valid user exists
        rsRecord.MoveFirst
        Idx = 1
        Do While Not rsRecord.EOF
            pUser.USER = rsRecord("UserName")
            pUser.PWord = rsRecord("PassCode")
            pUser.Access = rsRecord("Access")
            If UCase(pUser.USER) <> "APS" Then
                SelectUserName.AddItem pUser.USER, Idx
                Idx = Idx + 1
            End If
            rsRecord.MoveNext
        Loop
    End If
    rsRecord.Close
    dbDbase.Close

End Sub

Private Sub Form_Load()
Dim flag As Boolean
Dim hilim As Single
Dim iStn As Integer
    KeyPreview = True
    ' Set Title Foreground color
    lblPidControl(1).ForeColor = Titles_ForeColor
    lblPidControl(2).ForeColor = Titles_ForeColor
    ' set message forecolor
    lblMessage.ForeColor = Message_ForeColor
    lblMessage.Caption = vbCrLf & "Current Configuration Settings"
    cmdPrint.Visible = IIf(PRINTERAVAILABLE, True, False)
    ' Butane Options
    If systemhasBUTANE Then
        frmOotResp(ootBtnFlow).Visible = True
        lblButFlowTol.Visible = True
        lblButFlowUnits.Visible = True
        txtBtnTol.Visible = True
        lblMixTol.Visible = True
        lblMixUnits.Visible = True
        txtMixRatio.Visible = True
        lblNitFlowTol.Visible = True
        lblNitFlowUnits.Visible = True
        txtNitTol.Visible = True
        If systemhasORVR2 Then
            lblORVRButFlowTol.Visible = True
            lblORVRButFlowUnits.Visible = True
            txtORVRBtnTol.Visible = True
            lblORVRNitFlowTol.Visible = True
            lblORVRNitFlowUnits.Visible = True
            txtORVRNitTol.Visible = True
        Else
            lblORVRButFlowTol.Visible = False
            lblORVRButFlowUnits.Visible = False
            txtORVRBtnTol.Visible = False
            lblORVRNitFlowTol.Visible = False
            lblORVRNitFlowUnits.Visible = False
            txtORVRNitTol.Visible = False
        End If
    Else
        frmOotResp(ootBtnFlow).Visible = False
        lblButFlowTol.Visible = False
        lblButFlowUnits.Visible = False
        txtBtnTol.Visible = False
        lblMixTol.Visible = False
        lblMixUnits.Visible = False
        txtMixRatio.Visible = False
        lblNitFlowTol.Visible = False
        lblNitFlowUnits.Visible = False
        txtNitTol.Visible = False
        lblORVRButFlowTol.Visible = False
        lblORVRButFlowUnits.Visible = False
        txtORVRBtnTol.Visible = False
        lblORVRNitFlowTol.Visible = False
        lblORVRNitFlowUnits.Visible = False
        txtORVRNitTol.Visible = False
    End If
    
    If USINGREMSTSMON Then
        lblRemStatusLogInterval.Visible = True
        lblRemStatusLogIntervalUnits.Visible = True
        txtRemStatusLogInterval.Visible = True
        txtRemStatusLogInterval.Enabled = True
    Else
        lblRemStatusLogInterval.Visible = False
        lblRemStatusLogIntervalUnits.Visible = False
        txtRemStatusLogInterval.Visible = False
        txtRemStatusLogInterval.text = "10"
    End If
    
    ' Out-Of-Tolerance Response
    ResponseOOT(ootBtnFlow).ListIndex = SysConfig.BtnFlowResp - 1
    ResponseOOT(ootNitFlow).ListIndex = SysConfig.NitFlowResp - 1
    ResponseOOT(ootFuelTemp).ListIndex = SysConfig.FuelTempResp - 1
    ResponseOOT(ootFuelLevel).ListIndex = SysConfig.FuelLevelResp - 1
    ResponseOOT(ootStorageLevel).ListIndex = SysConfig.StorageLevelResp - 1
    ResponseOOT(ootPurFlow).ListIndex = SysConfig.PurFlowResp - 1
    ResponseOOT(ootAirMoist).ListIndex = SysConfig.AirMoistResp - 1
    ResponseOOT(ootAirTemp).ListIndex = SysConfig.AirTempResp - 1
    ResponseOOT(ootCanVent).ListIndex = SysConfig.CanVentResp - 1
    ResponseOOT(ootLoadRate).ListIndex = SysConfig.LoadRateResp - 1
    ResponseOOT(ootPurgeDp).ListIndex = SysConfig.PurgeDpResp - 1
    ResponseOOT(ootPurgeOvenTemp).ListIndex = SysConfig.PurgeOvenResp - 1
    ResponseOOT(ootWaterBathTemp).ListIndex = SysConfig.WaterBathResp - 1
        
    
    ' LiveFuel Options
    LiveFuelStn = 0
    NumberOfLiveFuelStations = 0
    If systemhasLIVEFUEL Then
        For iStn = 1 To LAST_STN
            If ((STN_INFO(iStn).Type = STN_LIVEFUEL_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEREG_TYPE) Or (STN_INFO(iStn).Type = STN_LIVEORVR2_TYPE)) Then
                If NumberOfLiveFuelStations = 0 Then LiveFuelStn = iStn
                NumberOfLiveFuelStations = NumberOfLiveFuelStations + 1
            End If
        Next iStn
        ResponseOOT(ootNitFlow).ToolTipText = "Select Response to Nitrogen(Vapor Carrier) Flow Rate OutOfTolerance condition"
    Else
        ResponseOOT(ootNitFlow).ToolTipText = "Select Response to Nitrogen Flow Rate OutOfTolerance condition"
    End If
    If systemhasLIVEFUEL And LiveFuelStn > 0 And LiveFuelStn <= LAST_STN Then
        txtDispStn.text = Format(LiveFuelStn, "#0")
        lblFuelFlowTol.Visible = True
        lblFuelFlowUnits.Visible = True
        txtLfvTol.Visible = True
        frmOotResp(ootLoadRate).Visible = True
        If systemhasAUTODRAINFILL Then
            If STN_INFO(iStn).ADF_DEF.hasADF_FST Then
                frmOotResp(ootStorageLevel).Visible = True
            Else
                frmOotResp(ootStorageLevel).Visible = False
            End If
            frmOotResp(ootFuelLevel).Visible = True
            frmOotResp(ootFuelTemp).Visible = True
            cfgtabs.TabVisible(Tab_AutoDrainFill) = True
            Disp_ADF LiveFuelStn
        Else
            frmOotResp(ootStorageLevel).Visible = False
            frmOotResp(ootFuelLevel).Visible = False
            frmOotResp(ootFuelTemp).Visible = False
            cfgtabs.TabVisible(Tab_AutoDrainFill) = False
            txtFuelTempTol.Visible = False
            lblFuelTempTol.Visible = False
            lblFuelTempUnits.Visible = False
        End If
    Else
        frmOotResp(ootStorageLevel).Visible = False
        frmOotResp(ootFuelLevel).Visible = False
        frmOotResp(ootFuelTemp).Visible = False
        frmOotResp(ootLoadRate).Visible = False
        cfgtabs.TabVisible(Tab_AutoDrainFill) = False
        lblFuelFlowTol.Visible = False
        lblFuelFlowUnits.Visible = False
        txtLfvTol.Visible = False
        txtFuelTempTol.Visible = False
        lblFuelTempTol.Visible = False
        lblFuelTempUnits.Visible = False
        txtVaporGenLeakRate.Visible = False
        lblVaporGenLeakRate.Visible = False
        lblVaporGenLeakRateUnits.Visible = False
        txtFuelStorageLeakRate.Visible = False
        lblFuelStorageLeakRate.Visible = False
        lblFuelStorageLeakRateUnits.Visible = False
    End If
    
    
    flag = IIf(USINGLOADPRESSURE, True, False)
       lblLoadPressure.Enabled = flag
       lblLoadPressure.Visible = flag
       lblLoadPressureUnits.Enabled = flag
       lblLoadPressureUnits.Visible = flag
       txtLoadPressure.Enabled = flag
       txtLoadPressure.Visible = flag
    
    flag = IIf(USINGBUTANEMASSLIMIT, True, False)
       lblButaneMassLimit.Enabled = flag
       lblButaneMassLimit.Visible = flag
       lblButaneMassLimitUnits.Enabled = flag
       lblButaneMassLimitUnits.Visible = flag
       txtButaneMassLimit.Enabled = flag
       txtButaneMassLimit.Visible = flag
    
    flag = IIf(USINGLOADTIMELIMIT, True, False)
       lblLoadTimeLimit.Enabled = flag
       lblLoadTimeLimit.Visible = flag
       lblLoadTimeLimitUnits.Enabled = flag
       lblLoadTimeLimitUnits.Visible = flag
       txtLoadTimeLimit.Enabled = flag
       txtLoadTimeLimit.Visible = flag
    
    optDbfBackup.Enabled = True
    optDbfBackup.Visible = True
    txtDbfBackupPath.Enabled = True
    txtDbfBackupPath.Visible = True
    
    optRptBackup.Enabled = True
    optRptBackup.Visible = True
    txtRptBackupPath.Enabled = True
    txtRptBackupPath.Visible = True
    
    If USINGCONTAFTERLCFAIL Then
        lblLeakErrResponse.Visible = True
        SelectLkChkFailResp.Enabled = True
        SelectLkChkFailResp.Visible = True
    Else
        ' no choice; default = STOP on Leak Check Failure
        lblLeakErrResponse.Visible = False
        SelectLkChkFailResp.Enabled = False
        SelectLkChkFailResp.Visible = False
        SysConfig.LeakCheckFailResponse = 0
    End If
    
    If USINGPRESSUREPURGE Then
        chkPosPressPurge.Enabled = True
        chkPosPressPurge.Visible = True
    Else
        chkPosPressPurge.Enabled = False
        chkPosPressPurge.Visible = False
        chkPosPressPurge.Value = cNO
    End If
    
    If USINGDRYPURGEAIR Then
        chkDryAirPurge.Enabled = True
        chkDryAirPurge.Visible = True
    Else
        chkDryAirPurge.Enabled = False
        chkDryAirPurge.Visible = False
        chkDryAirPurge.Value = cNO
    End If
    
    flag = IIf(USINGUPS > 0, True, False)
    lblUPSOpenDelay.Enabled = flag
    lblUPSOpenDelay.Visible = flag
    txtUPSOpenDelay.Enabled = flag
    txtUPSOpenDelay.Visible = flag
    
    flag = IIf(USINGDOOROPEN, True, False)
    lblDoorOpenDelay.Enabled = flag
    lblDoorOpenDelay.Visible = flag
    txtDoorOpenDelay.Enabled = flag
    txtDoorOpenDelay.Visible = flag
    
    flag = IIf(USINGOOTPAUSE, True, False)
    cfgtabs.TabEnabled(Tab_OotResponse) = flag
    
    flag = IIf((USINGCANVENTALARM), True, False)
    frmOotResp(ootCanVent).Visible = flag
    lblCanventDescr.Visible = flag
    lblCanventUnits.Visible = flag
    txtCanventOvr.Enabled = flag
    txtCanventOvr.Visible = flag
    
    flag = IIf((USINGCANVENTALARM Or systemhasAUTODRAINFILL), True, False)
    OotCol1Left = IIf(flag, 30, 960)
    OotCol2Left = IIf(flag, 3105, OutOfSight)
    OotCol3Left = IIf(flag, 6180, 5160)
    
    flag = IIf(USINGPURGEDP, True, False)
    frmOotResp(ootPurgeDp).Visible = flag
    lblPurgeDpHiLimit.Visible = flag
    lblPurgeDpHiLimitUnits.Visible = flag
    txtPurgeDpHiLimit.Enabled = flag
    txtPurgeDpHiLimit.Visible = flag
    
     flag = IIf(USINGPURGEOVEN, True, False)
    lblOvenBand.Visible = flag
    lblOvenBandUnits.Visible = flag
    txtOvenBand.Enabled = flag
    txtOvenBand.Visible = flag
    lblOvenTempTol.Visible = flag
    lblOvenTempUnits.Visible = flag
    txtOvenTempTol.Enabled = flag
    txtOvenTempTol.Visible = flag
    frmOotResp(ootPurgeOvenTemp).Visible = flag
    
     flag = IIf(USINGWATERBATH, True, False)
    lblWaterBathDesc.Visible = flag
    lblWaterBathUnits.Visible = flag
    txtWaterBathTol.Enabled = flag
    txtWaterBathTol.Visible = flag
    frmOotResp(ootWaterBathTemp).Visible = flag
    WaterBathTemperatureControl.Visible = flag
    
    ' align OOT Response boxes
    frmOotResp(ootBtnFlow).Left = OotCol1Left
    frmOotResp(ootNitFlow).Left = OotCol1Left
    frmOotResp(ootFuelTemp).Left = OotCol1Left
    frmOotResp(ootLoadRate).Left = OotCol1Left
    frmOotResp(ootWaterBathTemp).Left = OotCol1Left
    frmOotResp(ootCanVent).Left = OotCol2Left
    frmOotResp(ootPurFlow).Left = OotCol3Left
    frmOotResp(ootAirMoist).Left = OotCol3Left
    frmOotResp(ootAirTemp).Left = OotCol3Left
    frmOotResp(ootPurgeDp).Left = OotCol3Left
    frmOotResp(ootPurgeOvenTemp).Left = OotCol3Left
    
    ' Show & Allow Changes to Auto Logon Settings
    '   only if User has appropriate access code
    flag = IIf(CheckPass("Y", False), True, False)
    ' auto logon
    CfgAutoLogonInit
    Select Case AutoLogon
        Case autologonOFF
            ' Auto Logon is disabled
            lblAutoLogon.Visible = False
            SelectUserName.Visible = False
        Case autologonON
            ' Auto Logon is enabled
            lblAutoLogon.Visible = flag
            SelectUserName.Visible = flag
            lblAutoLogon.Enabled = True
            SelectUserName.Enabled = True
        Case autologonAPS
            ' Auto Logon is overridden (i.e. Logged on as Aps)
            lblAutoLogon.Visible = flag
            SelectUserName.Visible = flag
            lblAutoLogon.Enabled = False
            SelectUserName.Enabled = False
        Case Else
            ' disable Auto Logon
            AutoLogon = autologonOFF
            lblAutoLogon.Visible = False
            SelectUserName.Visible = False
    End Select
    
    flag = IIf((LocalPagControl.Type = pagClient) And CheckPass("C", False), False, True)
       lblTempTarget.Enabled = flag
       lblTempTargetUnits.Enabled = flag
       txtTempTarget.Enabled = flag
       lblMoistureTarget.Enabled = flag
       lblMoistTargetUnits.Enabled = flag
       txtMoistureTarget.Enabled = flag
       lblTempTol.Enabled = flag
       lblTempTolUnits.Enabled = flag
       txtTempTol.Enabled = flag
       lblMoistureTol.Enabled = flag
       lblMoistTolUnits.Enabled = flag
       txtMoistureTol.Enabled = flag
    
    ' PAS Local Control
    flag = IIf(USINGPASLOCALCONTROL And CheckPass("C", False), True, False)
    lblPidControl(pasTEMPERATURE).Visible = flag
    lblOnDutyMult(pasTEMPERATURE).Visible = flag
    txtOnDutyMult(pasTEMPERATURE).Visible = flag
    lblOffDutyMult(pasTEMPERATURE).Visible = flag
    txtOffDutyMult(pasTEMPERATURE).Visible = flag
    lblInTolDuration(pasTEMPERATURE).Visible = flag
    txtInTolDuration(pasTEMPERATURE).Visible = flag
    lblTimeoutDuration(pasTEMPERATURE).Visible = flag
    txtTimeoutDuration(pasTEMPERATURE).Visible = flag
    lblPidControl(pasMOISTURE).Visible = flag
    lblPgain(pasMOISTURE).Visible = flag
    txtPgain(pasMOISTURE).Visible = flag
    lblIgain(pasMOISTURE).Visible = flag
    txtIgain(pasMOISTURE).Visible = flag
    lblInTolDuration(pasMOISTURE).Visible = flag
    txtInTolDuration(pasMOISTURE).Visible = flag
    lblTimeoutDuration(pasMOISTURE).Visible = flag
    txtTimeoutDuration(pasMOISTURE).Visible = flag
    
    ' Temp & Humidity Logging
    flag = IIf(LogTempRh, True, False)
       lblLogTempRhInterval.Enabled = flag
       lblLogTempRhInterval.Visible = flag
       lblLogTempRhIntervalUnits.Enabled = flag
       lblLogTempRhIntervalUnits.Visible = flag
       txtLogTempRhInterval.Enabled = flag
       txtLogTempRhInterval.Visible = flag
       optLogTempRhVerbose.Enabled = flag
       optLogTempRhVerbose.Visible = flag
    
    ' set active tab
    frmConfig.cfgtabs.Tab = Tab_Job
    ' show Config screen
    Form_Center Me
    Update_Config   ' Update all configuration values with current values
End Sub

Private Sub chkDryAirPurge_Click()
    chkDryAirPurge.BackColor = cfgtabs.BackColor
    If (chkDryAirPurge.Value = cYES) Then chkPosPressPurge.Value = cNO
End Sub

Private Sub chkPosPressPurge_Click()
    chkPosPressPurge.BackColor = cfgtabs.BackColor
    If (chkPosPressPurge.Value = cYES) Then chkDryAirPurge.Value = cNO
End Sub

Private Sub txtBtnTol_Change()
    txtBtnTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtBtnTol_GotFocus()
    txtBtnTol.SelStart = 0
    txtBtnTol.SelLength = Len(txtBtnTol.text)
End Sub

Private Sub txtButaneMassLimit_Change()
    txtButaneMassLimit.BackColor = lblMessage.BackColor
End Sub

Private Sub txtCanventOvr_Change()
    txtCanventOvr.BackColor = lblMessage.BackColor
End Sub

Private Sub txtDoorOpenDelay_Change()
    txtDoorOpenDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtEventRecs_Change()
    txtEventRecs.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFileName_Change()
    txtFileName.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFileName_GotFocus()
    txtFileName.SelStart = 0
    txtFileName.SelLength = Len(txtFileName.text)
End Sub
Private Sub txtFileName_KeyPress(KeyAscii As Integer)
 If Not CheckPass("1", True) Then KeyAscii = 0
End Sub

Private Sub txtFuelStorageDrainDelay_Change()
    txtFuelStorageDrainDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageDrainShutoff_Change()
    txtFuelStorageDrainShutoff.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageDrainTimeout_Change()
    txtFuelStorageDrainTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageFillDelay_Change()
    txtFuelStorageFillDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageFillShutoff_Change()
    txtFuelStorageFillShutoff.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageFillTimeout_Change()
    txtFuelStorageFillTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelTempTol_Change()
    txtFuelTempTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtHeading_Change()
    txtHeading.BackColor = lblMessage.BackColor
End Sub

Private Sub txtHeading2_Change()
    txtHeading2.BackColor = lblMessage.BackColor
End Sub

Private Sub txtInTolDuration_Change(Index As Integer)
    txtInTolDuration(Index).BackColor = lblMessage.BackColor
End Sub

Private Sub txtJobRecs_Change()
    txtJobRecs.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLCIntvl_Change()
    txtLCIntvl.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLCMinDelay_Change()
    txtLCMinDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLCSetPoint_Change()
    txtLCSetPoint.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLCTime_Change()
    txtLCTime.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLfvTol_Change()
    txtLfvTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadIntvl_Change()
    txtLoadIntvl.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadIntvl_GotFocus()
    txtLoadIntvl.SelStart = 0
    txtLoadIntvl.SelLength = Len(txtLoadIntvl.text)
End Sub

Private Sub txtLoadPressure_Change()
    txtLoadPressure.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadTimeLimit_Change()
    txtLoadTimeLimit.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadTotal_Change()
    txtLoadTotal.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadTotal_GotFocus()
    txtLoadTotal.SelStart = 0
    txtLoadTotal.SelLength = Len(txtLoadTotal.text)
End Sub

Private Sub txtLoadTotIntvl_Change()
    txtLoadTotIntvl.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoLimLoad_Change()
    txtLoLimLoad.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLogTempRhInterval_Change()
    txtLogTempRhInterval.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoLimPurge_Change()
    txtLoLimPurge.BackColor = lblMessage.BackColor
End Sub

Private Sub txtMixRatio_Change()
    txtMixRatio.BackColor = lblMessage.BackColor
End Sub

Private Sub txtMixRatio_GotFocus()
    txtMixRatio.SelStart = 0
    txtMixRatio.SelLength = Len(txtMixRatio.text)
End Sub

Private Sub txtMoistureTarget_Change()
    txtMoistureTarget.BackColor = lblMessage.BackColor
End Sub

Private Sub txtMoistureTol_Change()
    txtMoistureTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtOffDutyMult_Change(Index As Integer)
    txtOffDutyMult(Index).BackColor = lblMessage.BackColor
End Sub

Private Sub txtOvenBand_Change()
    txtOvenBand.BackColor = lblMessage.BackColor
End Sub

Private Sub txtRemStatusLogInterval_Change()
    txtRemStatusLogInterval.BackColor = lblMessage.BackColor
End Sub

Private Sub txtRptBackupPath_Change()
    txtRptBackupPath.BackColor = lblMessage.BackColor
End Sub

Private Sub txtRptBackupPath_Click()
    frmBackupPath.ChangeBackupSelect SELECTRPTPATH
    frmBackupPath.Show
End Sub

Private Sub txtRptBackupPath_DblClick()
    frmBackupPath.ChangeBackupSelect SELECTRPTPATH
    frmBackupPath.Show
End Sub

Private Sub txtDbfBackupPath_Change()
    txtDbfBackupPath.BackColor = lblMessage.BackColor
End Sub

Private Sub txtDbfBackupPath_Click()
    frmBackupPath.ChangeBackupSelect SELECTDBFPATH
    frmBackupPath.Show
End Sub

Private Sub txtDbfBackupPath_DblClick()
    frmBackupPath.ChangeBackupSelect SELECTDBFPATH
    frmBackupPath.Show
End Sub

Private Sub txtNitrogenPurgeTime_Change()
    txtNitrogenPurgeTime.BackColor = lblMessage.BackColor
End Sub

Private Sub txtNitTol_Change()
    txtNitTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtNitTol_GotFocus()
    txtNitTol.SelStart = 0
    txtNitTol.SelLength = Len(txtNitTol.text)
End Sub

Private Sub txtOOTtime_Change()
    txtOOTtime.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPressureDecay_Change()
    txtPressureDecay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPurgeIntvl_Change()
    txtPurgeIntvl.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPurgeIntvl_GotFocus()
    txtPurgeIntvl.SelStart = 0
    txtPurgeIntvl.SelLength = Len(txtPurgeIntvl.text)
End Sub

Private Sub txtPurgeTol_Change()
    txtPurgeTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPurgeTol_GotFocus()
    txtPurgeTol.SelStart = 0
    txtPurgeTol.SelLength = Len(txtPurgeTol.text)
End Sub

Private Sub txtPurgeTotal_Change()
    txtPurgeTotal.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPurgeTotal_GotFocus()
    txtPurgeTotal.SelStart = 0
    txtPurgeTotal.SelLength = Len(txtPurgeTotal.text)
End Sub

Private Sub txtPurgeTotIntvl_Change()
    txtPurgeTotIntvl.BackColor = lblMessage.BackColor
End Sub

Private Sub txtTempTarget_Change()
    txtTempTarget.BackColor = lblMessage.BackColor
End Sub

Private Sub txtTempTarget_GotFocus()
    txtTempTarget.SelStart = 0
    txtTempTarget.SelLength = Len(txtTempTarget.text)
End Sub

Private Sub txtTempTol_Change()
    txtTempTol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtTempTol_GotFocus()
    txtTempTol.SelStart = 0
    txtTempTol.SelLength = Len(txtTempTol.text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    HotKeyCheck KeyCode, Shift  ' undo rest to display key coads
End Sub

Private Sub txtUPSOpenDelay_Change()
    txtUPSOpenDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLoadSettleTime_Change()
    txtLoadSettleTime.BackColor = lblMessage.BackColor
End Sub

Private Sub txtPurgeSettleTime_Change()
    txtPurgeSettleTime.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgDrainShutoff_Change()
    txtLiveFuelChgDrainShutoff.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgDrainDelay_Change()
    txtLiveFuelChgDrainDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgDrainTimeout_Change()
    txtLiveFuelChgDrainTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgFillDelay_Change()
    txtLiveFuelChgFillDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgFillTimeout_Change()
    txtLiveFuelChgFillTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgFillShutoff_Change()
    txtLiveFuelChgFillShutoff.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgHeaterTimeout_Change()
    txtLiveFuelChgHeaterTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgPurgeFillDelay_Change()
    txtLiveFuelChgPurgeFillDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgPurgeDrainDelay_Change()
    txtLiveFuelChgPurgeDrainDelay.BackColor = lblMessage.BackColor
End Sub

Private Sub txtLiveFuelChgPurgeTimeout_Change()
    txtLiveFuelChgPurgeTimeout.BackColor = lblMessage.BackColor
End Sub

Private Sub Print_Config()
' Procedure Name:   Print_Config
' Created By:       Brunrose
' Description:
'
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 53, 222
Dim iTab As Integer
Dim curTab As Integer
Dim maxTab As Integer
    
    ' Save current tab
    curTab = cfgtabs.Tab
    maxTab = IIf(USINGOOTPAUSE, 7, 6)
    
    ' CONFIG DATA
    Printer.Orientation = vbPRORLandscape
    ' cycle thru all the tabs
    For iTab = 0 To maxTab
            ' display the tab
            cfgtabs.Tab = iTab
            lblMessage.Caption = ""
            DoEvents
            ' capture the tab
            Set pbCapture.Picture = CaptureForm(Me)
            ' print the tab
            PrintPictureToFitPage Printer, pbCapture.Picture
            Printer.EndDoc
            ' short delay
            DoEvents
'            Delay_Box "", PAUSEDELAY, msgNOSHOW
    Next iTab
    ' clear capture box
    Set pbCapture.Picture = Nothing
    DoEvents

    ' Restore current tab
    cfgtabs.Tab = curTab
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

Private Sub txtFuelStorageLeakRate_Change()
    txtFuelStorageLeakRate.BackColor = lblMessage.BackColor
End Sub

Private Sub txtFuelStorageTankVol_Change()
    txtFuelStorageTankVol.BackColor = lblMessage.BackColor
End Sub

Private Sub txtVaporGenLeakRate_Change()
    txtVaporGenLeakRate.BackColor = lblMessage.BackColor
End Sub

Private Sub txtVaporGenTankVol_Change()
    txtVaporGenTankVol.BackColor = lblMessage.BackColor
End Sub

