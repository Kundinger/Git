Attribute VB_Name = "Module1"
' Error module 1 ''''''''''''''''''Program  GENERAL.bas '''''''''''''''''''''''''''''
' General Declarations Section
'*******************************************************************************
 'Change Log Release 2.0.1
'   Lux testing 5/14/2001 and shipped (All CPS code is ONE finally)
'      Purge by time added to all code
'   MarkIV June, 2001
'      First to have new code and all options Live Fuel, ORVR only, 2, 2 shift regular stations
'   Cosworth re-Compile 6/21/2001
'   Plastic  re-Compile 6/22/2001
'   New Bosch Add Station 2 Scale 2
'   Toyota recovery for problems 09/6/2001......Smitty .36 flow controllers again
'   AAA                          09/18/2001.....Smitty
'   Bosch upgrade to two stations 10/03/2001
'   New update for BMW 12/17/2001 Don't forget data base masters
'   Mark 4 recompile  1/30/2002   for Dave to take with
'   Bosch Recompile   2/25/2002
'   Mark 4 recompile  2/26/2002   for Dave/Smitty to take with
'   **  Fix I/O forcing causes scales in operation to end 2 gram breakthrough (06/21/02)
'   Mitsubishi (0376A) 9/9/2002) Using Mitsubishi to configure security default
'       Take this time to upgrade to newest version for purge time...etc.
'   NEW DATA BASE OPTIONS EVERYBODY FROM HERE ON NEEDS TO HAVE UPGRADE MASTER.....MDB
'   12/10/02 System testing back up stairs.....Smitty
'   12/13/02 Add pause for FID Expert .....Smitty
'   01/21/03 zero scales added     not fully tested yet...........................
'   CARB started 01/20/03
'   TJ FOUND A BAR NOT UPDATING IN 24 + HOUR LOAD BY TIME MODE   Thanks  1/24/03
'   Carb up stairs can vol purge done.....5/27/03 do load breakthrough now 5/28/03 down
'   5/30/03 up for report disable
'   6/06/03 Nitrogen push for usingcanvent
'   re-fix breakthrough & load volumes to match .....6/19/03
'   Special code fix OPTO's ....... sea level .......6/27/03
'   It was the scales getting alpha in the data stream; Not any more though (Coms modules) ..... 6/30/03 Smitty
'   Some nicities changes in read_station FIX piab restart logics 7/10/03 A through E
'   Fix OPTO configuration problems.... 7/24/03
'   Latest and greatest CARB1/2
'   Mark4 re-compile ....... 10/17/03 ....... Smitty
'   Mitsubishi Add new using option for abort based on timed load  Change "Z"  Smitty
'   FIX Mitsubishi Add new using option for abort based on timed load
'
'
'Change Log Release 3.0.1
'   NEW CHANGE CONTROL PROCEDURES IN EFFECT (ONLY ONE MASTER VERSION) - 1 Jan 2004 MMW
'
'   Fixed Live Fuel Option - Analog Output Value & Displayed PV....................................2   January 2004
'   Fixed File Maintenance - Explicitly exclude Event Log & Password Files........................22   January 2004
'
'   Added Live Fuel Chg Frequency Option  - Configurable # of Cycles Between Change Live Fuel Msg..1      June 2004
'
'   Added New Scale Type (A = Acculab) Option.....................................................28 September 2004
'   Changed "Nitrogen" to "TankAir Inlet" on Live Fuel Station Detail Form........................12   October 2004
'   Changed Live Fuel Summary & Detail Reports - TankAir Inlet vs. Nitrogen Flow..................15   October 2004
'
'   Added Auto LiveFuel Drain/Fill Option.........................................................22  November 2004
'   Manual LiveFuel Drain/Fill Form now appears without need to press Continue First..............23  November 2004
'   Added Optional Extra Cycle at 1.5 x Working Capacity Option...................................24  November 2004
'   Changed Displayed & Reported # of Cycles from 0 based to 1 based..............................25  November 2004
'   Added Auto LiveFuel Drain/Fill Mini-Display to Station Detail screen..........................20   January 2005
'
'   Fixed Problem with Load Cycles of 0.0 second Duration.........................................28   January 2005
'   Fixed Problem with Nitrogen & Butane Bar Graphs Failing to Update if Butane is Out of Range...28   January 2005
'   Modified Stop Button Operation to Stop Station even if a Station is Paused....................31   January 2005
'   Added Total Grams of Weight Loaded to the Summary Report for a Live Fuel station..............11  February 2005
'   Added "ORDER BY" to "Select Records from DB" criteria for Reports & Event Log form............14  February 2005
'   Removed Luxembourg Global Const & All Luxembourg Routines.....................................14  February 2005
'   Changed Report Filename convention;removed existing "." and added fixed extention of "RPT"....16  February 2005
'
'   Added USINGPRESSUREPURGE option...............................................................17  February 2005
'   Added ChgErrorModule routine..................................................................17  February 2005
'   Added Scale Aux Purge Air Valve support.......................................................25  February 2005
'   Renamed MainAspiratorORVR & MainAspORVR to MainPressPurgeSol & MainPressPurge..................4     March 2005
'   Added separate Global Const FILEPATH_'s for cal, cfg, data & reports...........................4     March 2005
'   Rewrote Purge Logic............................................................................5     March 2005
'   Revised Response to System Wide Alarms; Removed Alarm_Pausing(Vac, LEL, etc.)..................8     March 2005
'
'   Made use of MainAuxAir valves a compile option................................................14     March 2005
'   Removed Simulate compile option...............................................................14     March 2005
'   Removed Simulate & PLC routines...............................................................14     March 2005
'   Cleared Module2(Comms.bas) & removed frmTest..................................................14     March 2005
'   Added separate routine to control common Purge valves.........................................14     March 2005
'   Added Cycle Count Displays to the MainMenu....................................................15     March 2005
'
'   Tweaked colors on MainMenu for new computer...................................................18     April 2005
'
'   Removed check for LoadRate < 1 at LoadCycle_Start.............................................24     April 2005
'   Revised checks for LoadRate at StnRecipe Save.................................................24     April 2005
'   Fixed Unresponsive StationIO Forcing..........................................................11       May 2005
'   Fixed Ranges for Temp. %RH & Baro.............................................................11       May 2005
'
'   Removed "P Data" & "L Data" Errors............................................................17      June 2005
'   Revised "Questionable Data" Error Criteria....................................................17      June 2005
'
'   Tweaked I/O Config Form & FileCopy Logic......................................................10   October 2005
'
'   Added Dummy Stations and STN_INFO Array........................................................7  December 2005
'   Only Opto Reads are now from MainMenu Timer1..................................................11  December 2005
'   Added Multiple Station Auto LiveFuel Drain/Fill Option........................................11  December 2005
'   Added Individual Station Aspirators Option....................................................12  December 2005
'   Added Opt. 2nd Station Module to Stn IO Force Screen..........................................13  December 2005
'   Removed STN_HASORVR_AND_STANDARD as an option.................................................13  December 2005
'   Added ADF Heater Functionality................................................................17  December 2005
'   Removed USINGLOADPAUSE, USINGPURGEPAUSE and USINGPURGETIME as options.........................19  December 2005
'   Added ADF Drain/Fill Delay's & Timeout's to ConfigIO..........................................21  December 2005
'   Added ADF Heater Timeout to ConfigIO..........................................................28  December 2005
'   Rewrote ADF Sequence Routines & Removed frmAutoDrainFill screens...............................3   January 2006
'   Rewrote OOT_Check Routine......................................................................7   January 2006
'   Backcheck ADF level switches during drain & fill..............................................11   January 2006
'   Auto Drain waits for Sheath below 104 deg F...................................................12   January 2006
'   Mode in IO Force is set automatically.........................................................14   January 2006
'   Deleted routines from Module 11(Calib.bas) & Module2(comm.bas)................................15   January 2006
'   IO Address & Channel are now configured/defined (on SysDefStn)................................24   January 2006
'   Most Compile options moved to new SysDef.@@@ and changed to Boolean...........................24   January 2006
'   New Recipes.@@@, CanRcp.@@@, Scales.@@@ config files replace IOConfig2.@@@....................25   January 2006
'   All *.@@@ config files are now sized for MAX items and include many spares....................25   January 2006
'   Nitrogen Purge added to ADF Sequence & Heater Logic...........................................28   January 2006
'   Removed startup options "GetInputs" routine...................................................30   January 2006
'   New stations.@@@ & optoinfo.@@@ config files...................................................2  February 2006
'   New funcdefa.@@@ & funcdefd.@@@ config files...................................................6  February 2006
'   Removed GM_CETS everything.....................................................................8  February 2006
'   Allow Stopping from a Pause...................................................................10  February 2006
'   Added default values for Scales...............................................................13  February 2006
'   Made use of "Vac Sw on Purge Line" an option(normal=No Vac Sw.)...............................15  February 2006
'   Allow Automatic Heating with Manual Drain/Fill................................................16  February 2006
'
'   Fixed Overflow error on Continue..............................................................16   October 2006
'   Fixed PurgeInterlock Errors...................................................................24   October 2006
'   Modified Flow Totalization....................................................................25   October 2006
'
'   Opto Dll's to be copied to windows\system32....................................................5   January 2007
'   Added Delay before checking CanVentAlarm flow switch...........................................5   January 2007
'   Number of Recipes changed from 50 to 100......................................................16   January 2007
'   Number of Canister Recipes changed from 20 to 200.............................................16   January 2007
'   Added Display of "Delay before checking CanVentAlarm flow switch".............................22   January 2007
'   Fixed elapsed time calculation for time intervals that straddle midnight.......................7  February 2007
'   Data Writes for Load/Purge and "CheckStations" moved to separate MainMenu Timers..............14  February 2007
'   Added Station Sequence Controller.............................................................15  February 2007
'   Added "Purge Air Supply" Controller...........................................................16  February 2007
'   Added "Enter System Setup" option at Program Startup; No more compile options.................22  February 2007
'   Replaced separate Subroutine for each Comm Port with one arrayed Subroutine....................2     March 2007
'   New Timer on frmMainMenu for Reading Scales....................................................5     March 2007
'   Pulldown Menus for Scale Configuration.........................................................5     March 2007
'   Added "Purge Air Supply" Definition (CfgRevLvl 1)..............................................5     March 2007
'   Added Buttons to reset and set default Station and Function Definitions........................9     March 2007
'   Check (& correct) Calibration coefficients and MFC Range on Load...............................9     March 2007
'   ORVR & Regular Stations now differ in name only...............................................14     March 2007
'   Analog portion of IOForce Screen redesign.....................................................16     March 2007
'   Removed Analog data type......................................................................26     March 2007
'   Doors Open Too Long now Pauses System (used to Shut Down Program).............................30     April 2007
'   Read of a Comm Port is now reentrant..........................................................25       May 2007
'
'   Added check of Pri & Aux Scale indexes to "Load IOCfg".........................................9    August 2007
'   Butane Flow Tolerance Units is now Grams/Hour (was % of Full Scale)...........................22    August 2007
'   At beginning of LeakCheck, Open Purge Valve & MFC to vent existing pressure...................19 September 2007
'   Added MsComm errors to Error Handler..........................................................26 September 2007
'   Added (on config screen; 0.0-4.0%) LowLimit for Tolerance Checking of Purge & Load MFCs .......3   October 2007
'   Added SI/English Units choice (on sysdef) for LineVolume parameters (ID & Length) .............5   October 2007
'   Screen Colors are set by Windows Display Appearance settings..................................11   October 2007
'RELEASE 7.10.15 .................................................................................15   October 2007
'
'
'BEGIN RevLevel 2
'   Added optional environment variable WHEREISCPS to specify cps2000 root folder ................17   October 2007
'   Made MSGDELAY settable on sysdef main screen .................................................17   October 2007
'   DelayBox Routine now waits for the frmDelayBox to close ......................................17   October 2007
'   Added button on Calibration screen to set the ActualFlow values to nominal (FS only) .........19   October 2007
'   CPS2000 directory cleanup ....................................................................26   October 2007
'   Removed USINGTHERMOS ..........................................................................1  November 2007
'   Can now enter VDCmin & VDCmax for analog IO with non-integer values ...........................1  November 2007
'   Tidied up color references; only Global Const allowed; no hardcoded colors ...................16  November 2007
'   Program Revision Description is saved in sysdef (sysdef loaded then resaved on prog startup) .18  November 2007
'   Multiple scales assigned to the same comm port will all update ...............................18  November 2007
'   PurgeTime Calculations and error checking updated on StnRecipe & Recipe screens ..............18  November 2007
'   DB file cpsmodel.2007 replaces cpsmodel.2001 (adds millisecond resolution; CfgRevLvl 2) ......28  November 2007
'   Load & Purge Totalization intervals are now set on Config screen .............................29  November 2007
'   Added Station Mode Information Arrays ........................................................29  November 2007
'   Added System Timer monitor screen ............................................................29  November 2007
'   Cycle Count incremented after Load (or PauseAfterLoad if it is selected) .....................29  November 2007
'   Removed USINGSTARTLOADFIRST & USINGENDONPURGE (both options are now ALWAYS available) ........29  November 2007
'   Added GenerateReports choice to File Menu ....................................................29  November 2007
'   Added ReviewData choice to File Menu .........................................................29  November 2007
'
'BEGIN RevLevel 3
'   Master DB files renamed to include the rev level in the filename (DbfRevLvl 3) ...............29  November 2007
'   Config/sysdef files tweaked (CfgRevLvl 3) ....................................................29  November 2007
'   Config/sysdef file upgrade logic modified (CfgRevLvl 3) ......................................29  November 2007
'   Added check for existence of required files & folders at startup .............................29  November 2007
'   Tweaked access codes for Default User (no sysdef access; generate & review ok) ...............29  November 2007
'   Replaced Config screen DelayBox's with a Status panel & messages ..............................8   January 2008
'   Replaced Recipe screen DelayBox's with a Status panel & messages ..............................8   January 2008
'   Replaced StationRecipe screen DelayBox's with a Status panel & messages .......................8   January 2008
'   Corrected LoadRate & MixPercent checking during a StationRecipe save ..........................8   January 2008
'   Added WatchData choice to File Menu ...........................................................8   January 2008
'   Added (very simple) simulation option to sysdef ...............................................8   January 2008
'   Corrected bug in Scale InUse Logic; Scale ownership is now asserted before LeakWait ...........8   January 2008
'   Revised Main screen ...........................................................................8   January 2008
'   Corrected FileMaint_DB error that deleted open DB files that had not been written to yet) .....8   January 2008
'   Now write recipe & header data @ start of test; header updated @ end of test ..................8   January 2008
'   Manuals are now in PDF format; if only *.doc manuals exists, they will be used ...............22  February 2008
'   Added choices on Config screen for FileMaint and FileMaint_DB .................................3     March 2008
'   Added separate Vapor Carrier Flow Tolerance on Config screen ..................................8     March 2008
'   CPS2000 directory cleanup ....................................................................11     March 2008
'RELEASE 8.03.13 .................................................................................13     March 2008
'   update a - Corrected Station Recipe Save error with no Line Volume ...........................19     March 2008
'   update b - Corrected Station Recipe Save error with 0.0 Mix% and No Load .....................26     March 2008
'              Corrected wrong Purge Time on reports error .......................................26     March 2008
'   update c - Corrected Shift 2 not going into alarm ............................................31     March 2008
'   update d - Corrected display of Leak Check Press in PSIG on Common I/O Monitor................28     April 2008
'   update e - Added UPS Status Message to Main Menu..............................................12       May 2008
'
'
'BEGIN RevLevel 4
'   Added Delayed Start options to Recipes (CfgRevLvl 4) .........................................19     March 2008
'   Master DB files renamed to DbfRevLvl 4 .......................................................24     March 2008
'   Delayed Start Options added to Recipes Table in Model DB file (DbfRevLvl 4) ..................24     March 2008
'   Delayed Start Options added to Reports .......................................................24     March 2008
'   Automatically Continue (previously Idle stations only) after an Alarm clears .................26     March 2008
'   AutoLogon options added to sysdef ............................................................12       May 2008
'   Rearranged the File, Edit and View Menus .....................................................13       May 2008
'   Made "Description of External Alarm Contacts" configurable on sysdef screen ...................4      July 2008
'   Converted Process Recipes to a Custom Data Type ...............................................8      July 2008
'   Converted Canister Recipes to a Custom Data Type ..............................................9      July 2008
'   Combined Recipe and StnRecipe screens .........................................................9      July 2008
'   Combined CanRecipe and StnCanRecipe screens ...................................................9      July 2008
'
'BEGIN RevLevel 5
'   Replaced canrcp.@@@ with cpsrecipes db file (table = MasterCanisters)(CfgRevLvl 5) ............9      July 2008
'   Deleted Master CanRecipe array (Masters only in DB not in main memory anymore)(DbfRevLvl 5) ..10      July 2008
'   Added Copy and Paste options to the Canister Definition Screen ...............................10      July 2008
'   Replaced recipes.@@@ with cpsrecipes db file (table = MasterRecipes) .........................11 September 2008
'   Deleted Master Recipe array (Masters only in DB not in main memory anymore) ..................11 September 2008
'   For CfgRevLvl < 5, copy Canisters & Recipes from @@@ files to cpsrecipes db ..................11 September 2008
'   Calibration.xls file moved from the CPS2000 folder to the CPS2000\calibrate folder ...........11 September 2008
'   Added Screen to search (& sort & select) recipes .............................................11 September 2008
'   Added Screen to search (& sort & select) canisters ...........................................11 September 2008
'   Added New db file cpsSysDef.mdb in a new sysdbf folder........................................11 September 2008
'   Moved cpsMaster & cpsUser db files to the new sysdbf folder...................................11 September 2008
'   IOCfg.@@@ is no longer in use; new tables in cpsSysDef are used ..............................11 September 2008
'   Added sorting to Job List ....................................................................11 September 2008
'   Converted Configuration Data to a Custom Data Type ...........................................16 September 2008
'   Added data recording during Leak Check; data shown on Detail report ..........................16 September 2008
'   Renamed ErrorLog to EventLog .................................................................18 September 2008
'   Logon screen is shown at startup if no AutoLogon .............................................19 September 2008
'   At Test start, station gets private copy of Configuration Data (also save Cfg Data to dbf) ...20 September 2008
'   When printing reports, use the config values saved in the dbf ................................20 September 2008
'   Renamed Network options to Backup options ....................................................20 September 2008
'   Added Report Filename to Job List and Header data tables .....................................22 September 2008
'   Added Generate Reports, Display Reports & Copy Job Files to Joblist screen ...................22 September 2008
'   Added Leak Check data watching to the DataWatcher screen .....................................22 September 2008
'   Added Pause after Leak Check .................................................................23 September 2008
'   Startup Messages now show on the About screen ................................................24 September 2008
'   Removed 'Detail' from View Menu; Only show 'System Definition'(Edit menu) if User has access .25 September 2008
'   Added Menus and Toolbars to the MainMenu screen ...............................................2   October 2008
'   Added Menus and Toolbars to the Joblist screen ................................................7   October 2008
'   Application renamed to CPS_r5 (new default root folder = c:\cps_r5) ..........................10   October 2008
'   Added optional (config) Operator entered Report Filename .....................................12   October 2008
'   Common statusbars & message panels implemented ...............................................13   October 2008
'   Added Local PAS Control of Temp & Moisture ...................................................22   October 2008
'   Redesigned Config screen using tabs ...........................................................2  November 2008
'   Added ORVR2 station type (2 MFC's per Load Gas) ...............................................5  November 2008
'   Calibration is set to Read-Only when screen is first loaded ..................................18  November 2008
'   AutoPrint functionality restored; also Report printing from the JobList screen................18  November 2008
'   Added dynamic ToolTips to all elements of the system status bar...............................20  November 2008
'   Purge MFC gets SP after the respective PIAB sol turns on ......................................4  December 2008
'   Added separate OOT settings for each ORVR2 MFC ................................................4  December 2008
'   Expanded PurgeAir Monitor Screen (now available whenever logged in as ApsUser) ................7   January 2009
'   vbIgnore after Error 3420 in frmReview.JobComplete; error written to Elog, no msg displayed ..15   January 2009
'RELEASE 2009.01.20 ..............................................................................20   January 2009
'            - Corrected "Object Required" error on startup.......................................20   January 2009
'   update a                                            released 29 May 2009
'            - Corrected LeakCheck Pressure Col. Header on Watch & Review screens.................21   January 2009
'            - Corrected error in reading of TC inputs (note EUmin/max should be 0/100)...........21   January 2009
'            - Added delay between load valves & load mfc's........................................4  February 2009
'            - Statistics now updated only when data is logged.....................................5  February 2009
'            - Removed USINGREMOTEVALVES and frmstnValveControl....................................9  February 2009
'            - Added noise to simulated values.....................................................9  February 2009
'            - Added Purge-By-Time and Purge-Aux-Can-Only.........................................11  February 2009
'            - Added Live Fuel ADF Tank simulation.................................................4     March 2009
'            - Restored CopyFiles & PrintFiles screens and menu choices...........................16     March 2009
'            - Corrected error logging LeakCheck data.............................................25     March 2009
'            - Added zLog - PurgeLog including cpsZlog db file....................................30     March 2009
'            - Added Operator Manual button to navigation toolbar.................................17     April 2009
'            - Added FirstAid-File-Save-for-APS to menus..........................................23     April 2009
'            - Added Concordance Display to Station Detail screen (APS only).......................3       May 2009
'   update b                                            released 29 May 2009
'            - Corrected PAS check error (not checking Temp & Moist when not purging).............20      June 2009
'            - Corrected "back door" access to setup at startup...................................20      June 2009
'            - Corrected LowButane Overflow error.................................................26      June 2009
'            - LowButane Warning Status Change & Save of Settings by the Operator logged in Elog..26      June 2009
'            - Changed Error Module & Level stack from 20 to 100..................................30      June 2009
'            - Bypass-Program-Error-Msgs choice on sysdef now operational..........................7      July 2009
'   update c                                            released 8 July 2009
'            - Added Butane Density Multiplier for every Butane MFC................................8      July 2009
'   update d                                            released 11 July 2009
'            -1 Corrected Open-Station-Detail-screen-from-MainMenu-StnPanel error.................26      July 2009
'            -2 Corrected Missing-Last-Line-of-Detail-Report-Load-Detail-Data error...............10    August 2009
'            -3 Corrected Purge Time units in Recipe section of reports...........................14    August 2009
'            -4 Target(& Actual) for PurgeByVolume is now in Canister Volumes, not Liters.........19    August 2009
'            -5 Added check for spurious scale weight readings....................................19    August 2009
'            -6 Added revision sublevel to db filenames (CpsMaster_rev0500.mdb)....................9 September 2009
'            -6 Added revision sublevel to db filenames (CpsModel_rev0500.mdb).....................9 September 2009
'            -6 Added revision sublevel to db filenames (CpsSysDef_rev0500.mdb)....................9 September 2009
'            -6 Added revision sublevel to db filenames (CpsUser_rev0500.mdb)......................9 September 2009
'            -6 Added revision sublevel to db filenames (CpsZlog_rev0500.mdb)......................9 September 2009
'            -6 Added revision sublevel to db filenames (CpsRecipes_rev0500.mdb)...................9 September 2009
'            -6 Added debug logging to zLog of scale readings (CpsZlog_rev0501.mdb)................9 September 2009
'            -6 Added Load Settle Time to Recipes (CpsRecipes_rev0501.mdb & CpsModel_rev0501.mdb)..9 September 2009
'            -6 Added logging of data for running stations not in Load, Purge or Leakcheck.........9 September 2009
'            -7 Allow decimal minutes for Recipe Pause after Leak (CpsModel_rev0501.mdb)..........11 September 2009
'            -7 Allow decimal minutes for Recipe Pause after Load (CpsModel_rev0501.mdb)..........11 September 2009
'            -7 Allow decimal minutes for Recipe Pause after Purge (CpsModel_rev0501.mdb).........11 September 2009
'            -7 Allow decimal minutes for Recipe Load Settle Time (CpsModel_rev0501.mdb)..........11 September 2009
'            -7 Added XY Charting of Net Grams of Butane on Station Detail screen.................11 September 2009
'            -8 Corrected x-axis on XY Charting of station Net Grams of Butane....................16 September 2009
'            -9 Only do OOT check for Load SP = 0 during LoadLoading..............................16 September 2009
'            -a Added Estimated Job Duration to the Station Detail screen.........................17 September 2009
'            -b "Beta" debugging version..........................................................17 September 2009
'   update e                                            released 18 September 2009
'            - Corrected error when closing Cal & Calcheck screens................................21 September 2009
'            - Moved Load Settle Time to Config (CpsModel_rev0502.mdb)............................21 September 2009
'            - Added Purge Settle Time to Config (CpsModel_rev0502.mdb)...........................21 September 2009
'            - Only do OOT check for Purge SP = 0 during PurgePurging.............................21 September 2009
'            - Corrected Load Statistics; Only calculate load stats during LoadLoading............21 September 2009
'            - Corrected Purge Statistics; Only calculate purge stats during PurgePurging.........21 September 2009
'            - Tweaked Estimated Job Duration calculation.........................................21 September 2009
'            - Corrected XYchart error when a station that is not the displayed station finishes..22 September 2009
'            - Changed XYchart "Primary Scale" pen color to Blue..................................22 September 2009
'            - XYchart now cleared on Job Start...................................................22 September 2009
'            - Corrected XY Graph no-update-error when opening StnDetail from other screens.......23 September 2009
'            - Corrected Purge Total line on Summary Reports (Purge-By-Vol only)..................23 September 2009
'            - Corrected No-Canister-&-Recipe-Buttons on Stn Detail when AutoLogon not used.......23 September 2009
'            - aps, APS & ApsUser have Access to SysDef; Access is deleted for other Usernames....23 September 2009
'   update f                                            released 23 September 2009
'            -1 Added check of Purge Rate when a Recipe is saved to a station......................5   October 2009
'            -1 Initialize station Pri & Aux scale readings (strings) to "0.00"....................5   October 2009
'            -1 Added Job Sequence Courses (CpsRecipes_rev0502.mdb)................................5   October 2009
'            -1 Added new DSN "CpsRecipes" = CpsRecipes_rev0502.mdb................................5   October 2009
'            -1 AutoLogon options added to Config (CpsModel_rev0503.mdb)...........................6   October 2009
'            -1 System will now wait for PAS Ready before resuming Purge from OOT .................8   October 2009
'            -1 Added simulation error for PAS Temp & RH (CpsSysdef_rev0501.mdb)...................8   October 2009
'            -1 Removed users APS & aps from User DB; now hardcoded (CpsUser_rev0501.mdb).........13   October 2009
'            -1 Removed duplicate entries from User DB; CASE doesn't matter (CpsUser_rev0501.mdb).13   October 2009
'            -1 Added debug logging to zLog of local PAS control values (CpsZlog_rev0502.mdb).....14   October 2009
'            -1 Added Calibration DB File (CpsCalibrations_rev0500.mdb)...........................22   October 2009
'            -2 Added new Modes for "Wait" & "Pause" Job Sequence Courses ........................10  November 2009
'            -2 Corrected Elog message when Auto_Stop on LeakCheck Error .........................10  November 2009
'            -2 Butane Supply values are now saved to the Windows Registry .......................13  November 2009
'   update g                                            released 13  November 2009
'            _1 Corrected Baro units on reports  (Barometer is always mBar)........................3  February 2010
'            _2 Corrected LeakCheck Pressure units on reports & displays (Pressure is always psig).5  February 2010
'            _3 Corrected divide-by-zero error when during calc of recipe duration.................7      June 2010
'            _3 Corrected scale-has-errors when simulating scales that are not currently in use....7      June 2010
'            _3 Corrected scroll bars and excessive width of Event Log & Data Log screens..........7      June 2010
'            _4 During Purge, "Actual" only updated during PurgePurging...........................26    August 2010
'            _4 LoadByWeight & LoadByBreakthru now have a minimum duration (LoadMinDuration(stn)).26    August 2010
'            _4 After EndOnPurge Purge, increment cycle number....................................26    August 2010
'            _4 Tweaked appearance of the EventLog & DataLog(Alarm,FileMaint,OOT) screens.........26    August 2010
'            _4 Program-Error-Notification on status bar is cleared when EventLog loads...........30    August 2010
'            _4 Corrected Reporting of Butane Total on reports when UsingLineVolume ..............30    August 2010
'            _5 Changed MAXVALIDWEIGHTCHANGE to 0.55 (was 0.25) ..................................25   October 2010
'            _5 Tweaked appearance of the EventLog & DataLog(Alarm,FileMaint,OOT) screens.........25   October 2010
'            _6 Tweaked appearance of the Overview screen with 2Stn/1Shift........................17  December 2010
'            _6 Corrected Canvent Override Logic error when SystemTime (i.e. Now) has a glitch....18  December 2010
'            _6 Modified use of "ActiveTitleBar" color for use with Windows7......................20  December 2010
'            _6 Correct response to Close button on PrintSet screen...............................28  December 2010
'            _6 Changed MAXVALIDWEIGHTCHANGE to 0.85 (was 0.55) ..................................28  December 2010
'            _7 Added sysdef entry of Minimum Allowed Mfc SetPoint (% of fullscale)...............20   January 2011
'            _7 Butane Supply values are now saved to the Windows Registry once per hour .........21   January 2011
'            _8 Modifed Reporting of Scale Weight Change on Detail & Summary reports .............24     March 2011
'   update h                                            released 24  June 2011
'            _1 Changed MAXVALIDWEIGHTCHANGE to 0.99 (was 0.85) ..................................24      June 2011
'            _1 Added #-of-Days to StnDetail Elapsed Time display ................................24      June 2011
'            _1 Added option of Moisture in RH instead of grains/lb (CpsModel_rev0504.mdb) .......24      June 2011
'            _1 USINGOPPSELSCALES option deleted; is now always true..............................28      June 2011
'            _1 Combined various USING options into Integers in sysdef.@@@ .......................28      June 2011
'            _2 Show Concordance sooner for short load cycles .....................................1      July 2011
'            _2 Updated Printer Setup screen ......................................................1      July 2011
'            _2 Updated Canister & Recipe printing ................................................1      July 2011
'            _2 Updated Configuration printing ....................................................5      July 2011
'            _2 Tweaked Cal & CalCheck printing ...................................................5      July 2011
'            _2 Tweaked DataLog & EventLog printing ...............................................5      July 2011
'            _2 Tweaked JobsList printing .........................................................5      July 2011
'            _2 Concordance available to CPS, ADMIN & APS (but not USER) ..........................6      July 2011
'            _2 Changed icon for Overview screen ..................................................8    August 2011
'            _3 Fixed display of Station Name on Data Watcher screen .............................15 September 2011
'            _4 Updated Common TC's ..............................................................11  December 2011
'            _4 BugFix: Not Read Addr=2 Analogs on Common Board ..................................23  December 2011
'            _5 Allow 16-slot Opto Boards for All Nodes ..........................................10     March 2012
'            _5 Allow No Opto Board for Station Nodes ............................................10     March 2012
'            _5 IoMonitor replaces ComIOForce, FIDIOForce & StnIOForce screens ...................19     March 2012
'            _5 BugFix: N2 Load SP when Butane & N2 MFC's have different ranges  .................12      June 2012
'   update j                                            released 12 June 2012
'            _1 BugFix: File_Maint is always enabled .............................................25      June 2012
'            _2 BugFix: For Aux Purge Only, a purge flow rate is not required .....................5 September 2012
'            _3 BugFix: With Aux Purge Only, no time value caused an error ........................5 September 2012
'            _4 Added Xls Report ..................................................................7     March 2013
'            _4 Expanded purge volume field to allow up to 9999 volumes (was 999) ................11     March 2013
'            _4 Added ButaneMassLimit & LoadTimeLimit to Configuration (CpsModel_rev0505.mdb) ....12     March 2013
'            _4 Added logic to Abort if Station ButaneMassLimit is exceeded ......................12     March 2013
'            _4 Added logic to Abort if Station LoadTimeLimit is exceeded ........................12     March 2013
'            _4 Added logic to Pause if Station LoadPressure Alarm ...............................12     March 2013
'            _4 Added EOT & Generation report options to Configuration (CpsModel_rev0506.mdb) ....21     March 2013
'   update k                                            released 21 March 2013
'            _1 BugFix: Weight Change Summary lines on reports corrected..........................25     March 2013
'            _1 BugFix: Baro field on Cal screen now non-zero ....................................25     March 2013
'            _1 BugFix: Load-End-Weight for Aux scale now before settling ........................25     March 2013
'            _2 BugFix: EOT Reporting generates DataSource error .................................26     March 2013
'            _3 BugFix: EOT Reporting TextEotSummary_AutoPrint generates runtime error 380 .......12      June 2013
'            _4 Added New Scale Type (A = A&D) Option ............................................14      June 2013
'            _5 Added Support for shift 3 & 4; (1-4 stations only) (CpsModel_rev0507.mdb).........30      July 2013
'            _6 Added Support for up to 4 aux. outputs per station (CpsRecipes_rev0503.mdb).......31      July 2013
'            _7 Added Display Properties Screen ...................................................2    August 2013
'            _7 Upgraded appearance of SysDef screens .............................................2    August 2013
'            _8 Added Opto Comm Off during Manual Report Generation ...............................4    August 2013
'            _8 Added Timer8 = (Manual) Report Generation .........................................9    August 2013
'            _8 ESTOP Input is now a SysDef choice ...............................................10    August 2013
'            _8 Added Scale Comm Off during Manual Report Generation .............................12    August 2013
'   update m                                            released 22 August 2013
'            _1 BugFix: Vehicle Number not shown in JobList until job is complete .................9 September 2013
'   update p                                            released 10 December 2013
'            _1 BugFix: On Main screen, "shift" area below pbar does not open StnDetail screen ...13  December 2013
'            _2 BugFix: USINGC, USINGLVol_Engl & USINGMoist_Grains were "Loaded" incorrectly  .....9   January 2014
'            _2 Changed min allowed LeakCheck duration to 10s from 15s ............................9   January 2014
'            _2 Now turn off Beacon when Operator presses Continue after PAUSE, StnAlarm or StnOOT 9   January 2014
'            _2 Added tabPB to Turn off Beacon ....................................................9   January 2014
'            _2 Added tabPB to Turn off Horn ......................................................9   January 2014
'            _3 BugFix: SearchCan was not filling the data grid ..................................30   January 2014
'            _3 BugFix: SearchRcp was not filling the data grid ..................................30   January 2014
'            _5 Added Temperature & Humidity Logging (AirLogs)....................................18    August 2014
'            _5 Added New db file in sysdbf folder for AirLogs (CpsAirLogModel_rev0501.mdb).......18    August 2014
'            _5 Added OOT checking of Air Temp & Humidity when using AirLogs .....................19    August 2014
'            _5 Added ViewAirLog screen to display contents of current AirLog file ...............19    August 2014
'            _5 Modified ViewAirLog screen to also display historical AirLog files ...............20    August 2014
'            _5 Added a trend chart to ViewAirLog screen .........................................21    August 2014
'   update q                                            released 22 August 2014
'            _1 BugFix: AnD Scale Value was not updating properly when "unstable".................15   October 2014
'   update r                                            released 15 October 2014
'            _1 BugFix: Changed flag for debug display of new scale reading to NotDebugSCALES ....16   October 2014
'            _1 OptionExplicit now on all Forms and Modules except 16 & 18 (from Opto22) .........16   October 2014
'            _2 Added verbose option to config for AirLog events to the EventLog ..................4  February 2015
'   update s                                            released 4 February 2015
'            _1 Added New db file in recipe folder for TOM Interface (CpsTomCanLoad_rev0502.mdb) ..9  February 2015
'            _1 Added TOM Interface screen and logic ..............................................9  February 2015
'            _1 Combined additional USING options into Integer#3 in sysdef.@@@ ...................10  February 2015
'            _1 Added TOM csv report .............................................................17  February 2015
'            _2 Added Clear(Reset) of Active TOM tasks at program startup ........................26  February 2015
'            _3 Changed ViewAirLog minimum log interval from 1 to 10 minutes ......................4     March 2015
'   update t                                            released 17 March 2015
'            _1 Added set "Vehicle" to the TOM VIN ...............................................19     March 2015
'            _2 Added Print Detail Report at End-Of-Test config option ...........................24     April 2015
'            _3 BugFix: Not Updating TOM DB properly ..............................................8       May 2015
'            _3 BugFix: Joblist screen GenerateReports never enabled if Dummy Station on system ..11       May 2015
'            _4 BugFix: Menu Item TomCanLoad didn't open TomCanLoad screen........................14       May 2015
'            _5 Added set "Engineer" to the TOM Specialist ........................................6      June 2015
'            _5 Added append TOM TaskID to the "Comment" ..........................................6      June 2015
'            _5 Removed TOM csv report ............................................................6      June 2015
'            _5 BugFix: Event Log now correctly records TOM Task job result........................6      June 2015
'   update v                                            not released
'            _1 Added 3 new station types (LiveFuel/Standard and 2x "future") ....................19     March 2015
'            _1 Added Support for Regular/LiveFuel Station Type (CpsRecipes_rev0504.mdb)..........19     March 2015
'            _1 Added new Leakcheck options(Primary Only, Aux Only or Both) to Recipe ............26     March 2015
'            _1 Added new Purge Method, Purge-By-Profile to Recipe ...............................27     March 2015
'            _1 Added new Purge Method, Purge-To-Target to Recipe ................................27     March 2015
'            _1 Added new Station DO's, Aux Direction Sol & Aux Leakcheck Sol ....................31     March 2015
'            _1 Added recipe changes to the Model DB (CpsModel_rev0508.mdb) ......................31     March 2015
'            _1 Added PurgeProfile screen and Purge Profile logic..................................1     April 2015
'            _1 Added Load, Purge and Station Control Blocks.......................................2     April 2015
'            _1 Added MasterProfiles & StationProfiles to CpsRecipes db...........................20     April 2015
'            _1 Added SearchProf screen ..........................................................22     April 2015
'            _1 Added PurgeProfile Progress Panel to StnDetail ...................................24     April 2015
'            _2 Added Print Detail Report at End-Of-Test config option (CpsModel_rev0510.mdb) ....24     April 2015
'            _2 Added OOT Control Block and separate response for each OOT to Configuration.......30     April 2015
'            _2 Upgraded Scale Simulation .........................................................7       May 2015
'            _3 Added SimulatedCanisterJobStart%Full to SimControlPanel ..........................11       May 2015
'            _3 Added NetWtChg to Purge Stats  ...................................................14       May 2015
'            _3 Modified PurgeProfile Progress Panel on StnDetail for Purge-To-Target ............15       May 2015
'            _4 Separated Report Generation from Cps_r5 Main program .............................15       May 2015
'            _4 Report Generation program started from JobList button ............................18       May 2015
'            _4 Report Generation code removed from Cps_r5 Main program ..........................21       May 2015
'            _4 Print Recipe(s) converted to WYSIWYG using PrintForm .............................21       May 2015
'            _4 Print Config converted to WYSIWYG using PrintForm ................................21       May 2015
'            _4 Print LiveFuel Config converted to WYSIWYG using PrintForm .......................21       May 2015
'            _4 Print FID Config converted to WYSIWYG using PrintForm ............................21       May 2015
'            _4 Added WYSIWYG PrintForm to SysdefMain screen .....................................21       May 2015
'            _4 Moved ReportGenerator EventsLog to Master (CpsMaster_rev0501.mdb) ................26       May 2015
'            _5 Job Data Writes now include Mode, Phase, LeakResult & ReportCode Descriptions ....28       May 2015
'            _5 Added Load, Purge, Start & End method Descriptions (CpsRecipes_rev0505.mdb).......29       May 2015
'
'BEGIN RevLevel 7
'            _0 Program name changed to Cps_r7.....................................................1      June 2015
'            _0 CPS & DB Rev Levels changed from 5 to 7 ...........................................1      June 2015
'            _0 Modified layout of StnDetail screen ...............................................1      June 2015
'            _0 Renamed Purge-To-Target to Purge-By-WorkingCapacity ...............................1      June 2015
'            _0 Added new Purge Method, Purge-To-Target to Recipe .................................1      June 2015
'            _0 Added new Purge Method, Purge-To-UndoLoad to Recipe ...............................2      June 2015
'            _0 Removed CurrentCondition Table from cpsSysDef_rev0704.mdb .........................9      June 2015
'            _0 Added LiveFuel Consumption Log to Master .........................................11      June 2015
'            _0 Moved LiveFuelTank display from StnDetail to (new)FuelSupply screen ..............17      June 2015
'            _0 Moved Concordance display from StnDetail to (new)Concordance screen ..............17      June 2015
'            _0 Added Fuel Storage Tank & Controls to FuelSupply screen ..........................18      June 2015
'            _0 Updated Model DB and Data Writes (esp for Live Fuel Changes) .....................22      June 2015
'            _0 Added LoadRatePID to Recipes (Recipes DB and Model DB) ...........................22      June 2015
'            _0 Cleanedup & integrated all WtChg calculations ....................................25      June 2015
'            _0 Added Drain & Fill Shutoff Levels to ADF Cfg (cpsSysDef_rev0704.mdb) ..............2      July 2015
'            _0 Added Sysdef Option "Hard-Piped Scales" ...........................................7      July 2015
'            _0 Added LeakCheckCanister(Description) to Model DB "Data" table .....................8      July 2015
'            _0 Added support for reporting of Pri & Aux LeakCheckResults .........................9      July 2015
'            _0 Modified reporting of LoadTotal Grams, Liters and WtChg (cpsModel_rev0711.mdb) ...10      July 2015
'            _0 Added Simulated LiveFuel density to SimControlPanel ..............................10      July 2015
'            _0 Modified Watchdog for Board0 for case where Beacon is not 0/13 ...................15      July 2015
'            _0 Added Course# to Data, Recipe, Stats (cpsModel_rev0712.mdb & cpsXlsReport.xls) ...17      July 2015
'            _0 Converted ConfigLiveFuel screen to new AutoDrainFill tab on Config screen ........27      July 2015
'            _0 Added Sequence & Courses Tables to Model DB (cpsModel_rev0714.mdb) ...............28      July 2015
'            _0 Reorganized VerticalBar, Oversize & Spawned boxes on StnDetail screen ............28      July 2015
'            _0 Updated HC Log screen; now displays/charts one month of data .....................29      July 2015
'            _0 Added Leakcheck Status Panel to StnDetail screen .................................31      July 2015
'RELEASE 7.01.03 .................................................................................31      July 2015
'            _0 BugFix:Courses (cpsXlsReport.xls, cpsModel_rev0715.mdb & cpsRecipes_rev0707.mdb) ..4    August 2015
'            _0 BugFix:Stop LiveFuel Load if Low Tank Level .......................................8    August 2015
'            _0 Added Line Volume parameters to Job Sequence .....................................11    August 2015
'            _0 BugFix:Response to OOT ...........................................................14    August 2015
'RELEASE 7.02.02 .................................................................................14    August 2015
'            _0 Added "SetDefaultSequence" to Job Sequence screen ................................18    August 2015
'            _0 Redesigned the JobSequence screen and logic (cpsRecipes_rev0708.mdb) .............24    August 2015
'            _0 BugFix:Selection of Purge Profile for Master Recipe ..............................27    August 2015
'            _0 Added End-Of-Test Xls Reporting options (cpsModel_rev0716.mdb) ...................27    August 2015
'RELEASE 7.02.04 .................................................................................27    August 2015
'            _0 Made system timer #8 (was ReportGen Timer) not enabled ...........................28    August 2015
'            _0 Added MAXINVALIDWEIGHTS to Sysdef Main screen (was a constant) ....................1 September 2015
'            _0 Added MAXVALIDWEIGHTCHANGE to Sysdef Main screen (was a constant) .................1 September 2015
'            _0 Changed VALIDWEIGHTCHANGE logic to a Running Average Calculation ..................1 September 2015
'            _0 Added:CycleType (cpsXlsReport.xls, cpsModel_rev0717.mdb & cpsRecipes_rev0709.mdb) .2 September 2015
'RELEASE 7.03.01 ..................................................................................2 September 2015
'            _0 Added support for EndWtChg EndMethod(cpsXlsReport.xls, cpsModel_rev0718.mdb) ......8 September 2015
'            _0 Added support for LiveFuel Tank(s) Volume in EU (cpsSysDef_rev0705.mdb) ..........10 September 2015
'            _0 Modified LiveFuel AutoDrainFill for Multiple Fuels in one Job Sequence ...........13 September 2015
'            _0 Modified StableWtChange EndMethod to meet Ford Standards(cpsModel_rev0719.mdb) ...17 September 2015
'            _0 Added Fuel Storage Drain&Fill Sequence and Config (cpsSysDef_rev0706.mdb) ........21 September 2015
'            _0 Added support for Multi-Course JobSequences to the Review screen .................22 September 2015
'            _0 Added navigation button for Fuel Consumption Log .................................23 September 2015
'RELEASE 7.03.03 .................................................................................23 September 2015
'            _0 Added support for EndCalcWc EndMethod(cpsXlsReport.xls) ..........................24 September 2015
'            _0 BugFix:FST Timeout disables ADF ..................................................28 September 2015
'            _0 FST Sequence now interlocked with Station Stop & Pause ...........................29 September 2015
'            _0 Added Job Event Log to Model DB(cpsModel_rev0720.mdb) ............................30 September 2015
'            _0 Added Calibration for Analog Inputs (cpsCalibrations_rev0701.mdb) .................2   October 2015
'            _0 Added Calibration for Scales (cpsCalibrations_rev0702.mdb) ........................9   October 2015
'            _0 Added Job Event Log to Xls Report (cpsXlsReport.xls) .............................14   October 2015
'            _0 Write All Recipes to JobDB at Beginning Of Job ...................................14   October 2015
'            _0 Added New Scale Type for Toledo Viper Scale ......................................15   October 2015
'            _0 Added Operator Pause .............................................................19   October 2015
'            _0 Added "View JobLog" button to JobList screen (cpsModel_rev0721.mdb) ..............19   October 2015
'            _0 Added SystemVacSw, PurgeDP & PurgeSeries to Sysdef ...............................20   October 2015
'            _0 Added SystemVacSw Common DI & PurgeDP Station AIs ................................20   October 2015
'            _0 Added Standard Temp & Press to all calibrations (cpsCalibrations_rev0703.mdb) ....20   October 2015
'            _0 Added Config & OOT checking for PurgeDP ..........................................21   October 2015
'            _0 Added Purge-Cans-In-Series to Recipe (cpsRecipes_rev0710.mdb) ....................21   October 2015
'            _0 Added Csv Reports to Config and to rev0721 Model DB ..............................23   October 2015
'            _0 Added SystemVacSw Alarm & Resume Logic ...........................................27   October 2015
'            _0 Added ValidCourse checks for CalcWC EndMethod .....................................9  November 2015
'            _0 Added Calibration Report for AnalogInput and Scale Calibrations ..................11  November 2015
'RELEASE 7.04.04 .................................................................................11  November 2015
'            _a BugFix:MFC Calibration doesn't calibrate .........................................12  November 2015
'            _b BugFix:Change operation of Aux Direction Valve ...................................18  November 2015
'            _c BugFix:"SIMULATED I/O..." in CommentBox when not in Simulation ...................24  November 2015
'            _c Added Display of Purge DP to StnDetail screen ....................................24  November 2015
'            _d BugFix:Valve Operation during LeakCheck and Purge(esp. Aux/Series Purge Valves) ...2  December 2015
'            _e BugFix:Purge Options Recipe selections ............................................3  December 2015
'            _e Modified Station Detail screen ....................................................8  December 2015
'            _e Added Purge DP to Job Data (cpsModel_rev0722.mdb & cpsXlsReport.xls) ..............9  December 2015
'            _e BugFix:Not showing negative Viper scale readings as negative .....................10  December 2015
'            _e BugFix:Not detecting loss of serial comm inputs ..................................14  December 2015
'            _e BugFix:Reported cycles when EndMethod=WeightChg (cpsModel_rev0723.mdb) ...........15  December 2015
'            _e Added PortValues screen ..........................................................24  December 2015
'            _e Removed support for CalcWc EndMethod (cpsXlsReport.xls) ..........................24  December 2015
'            _e Added support for CalcWc Recipe option (cpsRecipes_rev0711.mdb) ...................3   January 2016
'            _e Modified support for Update of Canister WC (cpsModel_rev0724.mdb) .................3   January 2016
'            _e Consistent Seq,Rcp,DB scale refs (cpsModel_rev0725.mdb, cpsRecipes_rev0712.mdb) ...6   January 2016
'            _e Added MasterMode & StationMode BackColors (cpsSysDef_rev0707.mdb) .................6   January 2016
'            _e Updated Calibration for MFCs (cpsCalibrations_rev0704.mdb) ........................7   January 2016
'            _e Added FuelLevel OOT (cpsModel_rev0726.mdb, cpsSysDef_rev0708.mdb) ................11   January 2016
'RELEASE 7.05.04 .................................................................................12   January 2016
'            _a Added StorageLevel OOT (cpsModel_rev0727.mdb, cpsSysDef_rev0709.mdb) .............22   January 2016
'            _a If default jobseq then saving recipe copies rcp duration to jobseq duration .......2  February 2016
'            _b Added new station type "Regular/LiveFuel with ORVR2" (CpsRecipes_rev0504.mdb) ....15     March 2016
'            _b Upgraded TomCanLoad to Remote Tasks & Monitor (cpsModel_rev0728.mdb) ..............3   January 2016
'            _b BugFix:Not starting ADF when required ............................................30      June 2016
'            _b Expanded allowed ADF Storage Tank Drain Timeout on Config screen to 3600sec .......7      July 2016
'            _c Added support for 8-slot Opto Board ..............................................25      July 2016
'            _c BugFix:Review Screen showed Course# when only 1 Sequence set on sysdef ...........28      July 2016
'            _c Added "Level Xmtr" to LiveFuel Options on sysdef stn screen ......................28      July 2016
'            _c Now:Ambient Air & PAS Temp & Humidity; was just PAS Temp & Humidity ..............28      July 2016
'            _d Now:FirstAid supports selection of multiple jobs .................................29      July 2016
'            _d BugFix:I/O Monitor showed first 8 slots as No Module .............................29      July 2016
'            _d BugFix:Tank Level & LoadRate OOT's now reset on Station Continue ..................2    August 2016
'            _d Added Sheath OverTemp Sw (DI) ....................................................11    August 2016
'            _d Revised ADF Sequence for Stant LiveFuel Tank; esp. N2 Pressurization .............11    August 2016
'            _d Added New MFC type - LiveFuel ORVR ...............................................11    August 2016
'            _d BugFix:PowerUpClear_300 for Analogs on node=base+2 ...............................12    August 2016
'            _d Added LiveFuel Tank Vent Sol......................................................12    August 2016
'            _d BugFix:CorrectedPurgeAir Panel Display ...........................................12    August 2016
'            _d Updated FuelSupply screen for No Fuel Storage Tank................................16    August 2016
'            _d Added PurgeWizard screen..........................................................17    August 2016
'            _d Added LiveFuel ORVR Solenoid......................................................19    August 2016
'            _d Added HardCoded LiveFuel Density..................................................19    August 2016
'            _d Added Clear & Delete buttons to SearchCan screen..................................22    August 2016
'            _d Added Clear & Delete buttons to SearchProf screen.................................22    August 2016
'            _d Added Clear & Delete buttons to SearchRcp screen..................................22    August 2016
'            _d Added Clear & Delete buttons to SearchJobSeq screen...............................22    August 2016
'            _d BugFix:LF ORVR MFC used for Hi-Range LoadRate PID.................................23    August 2016
'            _d Added Create New button to SearchCan screen.......................................23    August 2016
'            _d Added Create New button to SearchProf screen......................................23    August 2016
'            _d Added Create New button to SearchRcp screen.......................................23    August 2016
'            _d Added Create New button to SearchJobSeq screen....................................23    August 2016
'            _d Added Calc of Current LiveFuel Density............................................25    August 2016
'            _d Added Latest LeakCheckResult to StnDetail screen..................................25    August 2016
'            _d BugFix:LF MFC always logged even if Hi-Range LF Selected..........................25    August 2016
'            _d Added optional Copy of Mfc Cal to "shared" Mfc....................................26    August 2016
'            _d BugFix:Canister Volume on StnDetail was always zero...............................26    August 2016
'            _d Added Save Master Canister to RemoteDB............................................29    August 2016
'            _d Added Save Master Recipe to RemoteDB..............................................29    August 2016
'            _d LiveFuel `Tank Level OOT Check is now a sysdef choice..............................2 September 2016
'            _d Added Save Master PurgeProfiles to RemoteDB........................................4 September 2016
'            _d BugFix:Stn Recipe Paste overwrites (hardcoded) scale assignments...................4 September 2016
'            _d BugFix:After LC Error & Stn stop, don't turn off LC Exh Sol if another stn in LC...4 September 2016
'            _d Added "Reload Controller Parameters" button to PurgeMonitor screen.................6 September 2016
'            _d LiveFuel Max Sheath Temp for ADF Drain is now a sysdef value.......................6 September 2016
'            _d Added Refill LiveFuel Tank when Vapor Density (gm/liter) is too low................7 September 2016
'            _b AddedToSysdef: Min Allowed LiveFuel density (cpsModel_rev0729.mdb) ................7 September 2016
'            _b Added display of fuel state to FuelSupply screen ..................................8 September 2016
'            _d Added Refill LiveFuel Tank during Load when Vapor Density (gm/liter) is very low...9 September 2016
'            _d BugFix:After (Load w/ LF ORVR) ADF-Abort, didn't turn off LF ORVR Mfc .............9 September 2016
'            _d BugFix:Valves Not Energized for Mfc Cal or CalCheck ..............................13 September 2016
'            _d Added: Now using "Use Inverse" for DO's...........................................13 September 2016
'            _d BugFix:Recipe should be invalid if SeriesPurge is selected and PurgeAux is Not ...13 September 2016
'            _d Added: Import PurgeProfile from text file.........................................22 September 2016
'            _d AddedToSysdef: Dead & Weak LiveFuel densities (cpsModel_rev0730.mdb) .............10   October 2016
'            _d BugFix:MFC's not turned off when Stn is paused ...................................12   October 2016
'            _d BugFix:MFC's not turned off when PreLoad is Done .................................13   October 2016
'            _d Added Save Configuration to RemoteDB  ............................................18   October 2016
'            _d Added:Save Simulated LiveFuel Density (cpsSysDef_rev0710.mdb) ....................20   October 2016
'            _d BugFix:Added Canister description to Leakcheck messages in JobLog ................21   October 2016
'            _d Added Default-Report-Interval to Configuration  ..................................28   October 2016
'            _d Added support for Lauda Heater(Chiller) with serial command interface  ...........28      June 2017
'            _d Expanded support for Lauda Heater(Chiller)  (cpsRecipes_rev0713.mdb)..............24   October 2017
'            _d BugFix:EventLog was cleared at startup ............................................6  November 2017
'            _d Added support for Purge Oven  (cpsModel_rev0731.mdb, cpsRemote_rev0704.mdb)........6  November 2017
'            _d Added support for Dry Purge Air ...................................................9  November 2017
'            _d Added to Job DbFile: Dry Purge Air, PurgeOven, Chiller (cpsModel_rev0732.mdb) ....26  November 2017
'            _d Added to Job DbFile: Dry Purge Air, PurgeOven, Chiller (cpsRemote_rev0705.mdb) ...27  November 2017
'            _d BugFix:Didn't stop for manual Drain&Fill .........................................28  November 2017
'            _d BugFix:Recipe Not Valid if LeakCheck with no Primary Scale selected ..............28  November 2017
'            _d BugFix:No Piab during Cal(or CalCheck) of PurgeAir MFC ...........................28  November 2017
'RELEASE 7.05.06 .................................................................................29  November 2017
'            _a Added support for 40 CPR 1066.985 LeakTest Station  (cpsModel_rev0733.mdb)  .......6  December 2017
'            _b Added LeakTest sequence and LeakTest mode  (cpsModel_rev0734.mdb)  ...............15   January 2018
'            _b BugFix:Recipe - Purge Target weight allowed to be +/- 19999 ......................23   January 2018
'            _b Removed LeakTest icon from navigation buttons ....................................23   January 2018
'            _b Added LeakTest data logging  (cpsModel_rev0735.mdb)  .............................23   January 2018
'RELEASE 7.05.07 .................................................................................23   January 2018
'            _a BugFix:Couldn't Save an MFC CalCheck  (cpsCalibrations_rev0706.mdb)  ..............5     March 2018
'            _a Added MFC CalCheck History screen  ...............................................19     March 2018
'            _a Added RawValueType & Cal Method to AI & MFC Cal   ................................20     March 2018
'            _a Added Cal Help screen   ..........................................................22     March 2018
'RELEASE 7.05.08 .................................................................................22     March 2018
'            _a Added WaterBath Supervisory Temp Control  (cpsModel_rev0736.mdb)  .................4     April 2018
'            _a Added WaterBath Supervisory Temp Control  (cpsRemote_rev0706.mdb)  ...............11     April 2018
'            _a Added Out & CumI Max & Min to Controllers datatable (cpsSysDef_rev0711.mdb) ......11     April 2018
'            _a Added Out & CumI Max & Min to Controllers datatable (new DSN = cpsSysdef) ........11     April 2018
'            _a Added Added check for WB temp still OK at end of Manual Gas Pause   ..............11     April 2018
'            _a BugFix:On Recipe screen, Couldn't unselect a Pause after Load or Purge  ..........12     April 2018
'            _a BugFix:Error when Opening Empty Master Canisters, JobSeqs, Profiles or Recipes  ..20     April 2018
'            _a ADD (access "8" to admin in cpsUser_rev07xx.mdb)  ................................20     April 2018
'            _a BugFix:Never Unload frmComm8Card; just hide  .....................................23     April 2018
'            _a BugFix:Scale End Weights when Load or Purge is aborted  ..........................27     April 2018
'            _a BugFix:Use sysdef MfcSpMin for min SPs on Recipes (except where 0 is allowed)  ...27     April 2018
'            _a Removed File Maintenance  (cpsModel_rev0737.mdb)  .................................8       May 2018
'            _a BugFix: sysdef NrSeq not properly saved  .........................................12       May 2018
'            _a Added cps_r7 resource file  ......................................................17       May 2018
'            _a Removed frmAnalyser & frmConfigFID  ..............................................18       May 2018
'            _a Removed FID support (cpsModel_rev0738, cpsSysdef_rev0712) ........................19       May 2018
'            _a Removed FID support (cpsRemote_rev0708, cpsRecipes_rev0714) ......................19       May 2018
'            _a Revised Comm8 ErrorHandler  ......................................................20       May 2018
'            _a BugFix: Image for ViewJoblog on JobList screen ...................................20       May 2018
'            _a Removed cps_r7 resource file  .....................................................1      June 2018
'            _a Removed USINGREMOTECABINET  .......................................................1      June 2018
'            _a Removed STN_REMOTECABINET  ........................................................1      June 2018
'            _a BugFix: Concordance screen doesn't remain open ....................................1      June 2018
'            _a BugFix: NetWtChg screen doesn't remain open .......................................1      June 2018
'            _a BugFix: Owner of Pri Scale valves ................................................26      June 2018
'            _a BugFix: Load Aux-Scale-End-Weight when no Pri scale  ..............................5      July 2018
'            _a BugFix: Load Wt Chg Rate elapsed timer is reset when LoadLoading starts  .........13    August 2018
'            _a BugFix: Overflow of Integer counter in Load_Totalize .............................24    August 2018
'            _a Added Access version detection  ..................................................11 September 2018
'            _a BugFix: No ADF choice on Recipe screen for Tank Type < 11 ........................27 September 2018
'            _a Temporarily subverted Access version detection  ..................................28 September 2018
'            _a Added new Purge Method, Purge-By-Liters to Recipe (cpsRecipes_rev0715)............11   October 2018
'            _a Added new Purge Method, Purge-By-Liters to Recipe (cpsRemote_rev0709).............11   October 2018
'            _a Added new Purge Method, Purge-By-Liters to Recipe (cpsModel_rev0739)..............11   October 2018
'            _a BugFix: MFC description on Cal report ............................................15   October 2018
'RELEASE 7.07.01 ..................................................................................1  November 2018
'            _a BugFix: Nr of Recipes displayed on SysdefMain .....................................6  December 2018
'            _a Delete, Clear & CreateNew buttons on Search Rcp & Can screens only for Masters ....7  December 2018
'            _a BugFix: Duplicate Cum-Liters when Purge-By-Liters on Stn Detail....................7  December 2018
'            _a BugFix: Display of Stn# & Job# on DataWatcher .....................................7  December 2018
'            _a BugFix: Display of LeakTest data on DataWatcher ...................................7  December 2018
'            _a BugFix: Update of Remote SysdefMain Table ........................................15  February 2019
'            _a BugFix: No close box on Scale Cal screen .........................................11     April 2019
'RELEASE 7.07.02 .................................................................................24     April 2019
'            _a Added PAS AK Interface  (cpsZlog_rev0703; added AK_Log table).....................25     April 2019
'
'
'' Redesigned By:  Brunrose
'' Rewritten By:   Brunrose
'' Original Creation Date:  6/4/96
'
'
'
'     *** THIS VERSION ***
'     *** THIS VERSION ***
'     *** THIS VERSION ***
'Global Const VER_DESCRIPTION = "Release Version "           'Release Version
Global Const VER_DESCRIPTION = "Debug Version "                'Debug Version
Global Const PROGRAMREVLVL = "7.07.03"          ' Revision Level of Cps_r7 program
Global Const USINGRELEASEDATE = VER_DESCRIPTION & PROGRAMREVLVL & "  2019 April 25"   ' CHANGE DATE AS REQUIRED
Global Const RPTGENREVLVL = PROGRAMREVLVL       ' required revision level of ReportGenerator program
Global Const REMGENREVLVL = PROGRAMREVLVL       ' required revision level of Remote program
Global Const CPSREVLVL = 7                      ' master revision level for Config & Sysdef files
Global Const DBFREVLVL = 7                      ' master revision level for DB files
Global Const DBMODELREVLVL = 39                 ' revision sublevel for Model DB file
Global Const DBMASTERREVLVL = 2                 ' revision sublevel for Master DB file
Global Const DBCALREVLVL = 6                    ' revision sublevel for Calibration DB file
Global Const DBRCPREVLVL = 15                   ' revision sublevel for Recipe DB file
Global Const DBSYSDEFREVLVL = 12                ' revision sublevel for Sysdef DB file
Global Const DBUSERREVLVL = 1                   ' revision sublevel for User DB file
Global Const DBAIRLOGREVLVL = 1                 ' revision sublevel for AirLog DB file (temp & humidity)
Global Const DBREMREVLVL = 9                    ' revision sublevel for Remote Tasks DB file
Global Const DBZLOGREVLVL = 3                   ' revision sublevel for Zlog DB file
Global Const RESFILEREVLVL = 0                  ' revision sublevel for Resource file
'     *** THIS VERSION ***
'     *** THIS VERSION ***
'     *** THIS VERSION ***
'
'

' Constants for arrays
Global Const MAX_COMM = 16                      ' Max comm port number                      - For Array Sizing
Global Const MAX_PRG = 9                        ' Max number of PurgeAirSupplies            - For Array Sizing
Global Const MAX_SCALES = 16                    ' Max number of scales                      - For Array Sizing
Global Const MAX_SHIFT = 4                      ' Max number of shifts                      - For Array Sizing
Global Const MAX_STN = 9                        ' Max number of stations                    - For Array Sizing
Global Const MAX_CYCLES = 999                   ' Max number of test cycles                 - For Array Sizing
Global Const MAX_CANRCP = 200                   ' Max number of canister recipes            - For Array Sizing
Global Const MAX_RCP = 100                      ' Max number of test recipes                - For Array Sizing
Global Const MAX_PROFILES = 100                 ' Max number of purge profiles              - For Array Sizing
Global Const MAX_PROFILESTEPS = 2000            ' Max number of steps in a profile          - For Array Sizing
Global Const MAX_CONTROLLER = 30                ' Max number of PID or On/Off controllers   - For Array Sizing
Global Const MAX_COURSES = 99                   ' Max number of Courses in a Job Sequence   - For Array Sizing
' End constants for arrays

' Global Declarations
Global Const MAXCOMMERRORS = 9              ' # of Comm Port Errors before a Msg Box is Displayed
Global Const MAXWEIGHTQUEUE = 100           ' max # of elements in each scale weight readings queue
Global Const NumPoints = 1000               ' number of XY data points per chart pen
Global Const CALIBLIMIT = 0.25              ' Calibration Valid Limit
Global Const LiveFuelVaporDensity = 2.11    ' grams per liter (from: 95% => 5 slpm = 600 gm/hr, i.e. for 100%, 300 liters/hr = 632 gm/hr)

' Calibration Table Size Constants
Global Const MAXPINCALPOINTS = 12           ' maximum number of PinPoint calibration points
Global Const MINPINCALPOINTS = 3            ' minimum number of PinPoint calibration points
Global Const MAXLSQCALPOINTS = 11           ' The largest number of Least Squares calibration points
Global Const MINLSQCALPOINTS = 3            ' The smallest number of Least Squares calibration points
Global Const MAXCALCHECKS = 49              ' The maximum number of saved CalChecks per Calibration (AnalogInput or MFC or Scale)
' Constants to indicate the selected Units for the Raw Input Values for Analog Input Calibration
Global Const CalRawUndefined = 0
Global Const CalRawAsVolts = 1              ' Analog Input Entered Raw Values are in Volts
Global Const CalRawAsMa = 2                 ' Analog Input Entered Raw Values are in MilliAmperes
Global Const CalRawAsDegC = 3               ' Analog Input Entered Raw Values are in Degrees C
Global Const CalRawAsEU = 4                 ' Analog Input Entered Raw Values are in Engr Units
' Calibration AnalogInput Groups
Global Const calgrpComm = 0                  ' Common Analogs
Global Const calgrpStn1 = 1                  ' Station #1 Analogs
Global Const calgrpStn2 = 2                  ' Station #2 Analogs
Global Const calgrpStn3 = 3                  ' Station #3 Analogs
Global Const calgrpStn4 = 4                  ' Station #4 Analogs
Global Const calgrpStn5 = 5                  ' Station #5 Analogs
Global Const calgrpStn6 = 6                  ' Station #6 Analogs
Global Const calgrpStn7 = 7                  ' Station #7 Analogs
Global Const calgrpStn8 = 8                  ' Station #8 Analogs
Global Const calgrpStn9 = 9                  ' Station #9 Analogs
Global Const calgrpFid = 10                  ' FID Analogs
Global Const calgrpPrg1 = 11                 ' Purge #1 Analogs
Global Const calgrpPrg2 = 12                 ' Purge #2 Analogs
Global Const calgrpPrg3 = 13                 ' Purge #3 Analogs
Global Const calgrpPrg4 = 14                 ' Purge #4 Analogs
Global Const calgrpPrg5 = 15                 ' Purge #5 Analogs
Global Const calgrpPrg6 = 16                 ' Purge #6 Analogs
Global Const calgrpPrg7 = 17                 ' Purge #7 Analogs
Global Const calgrpPrg8 = 18                 ' Purge #8 Analogs
Global Const calgrpPrg9 = 19                 ' Purge #9 Analogs
' Calibration load constants
Global Const calNormal = 1                  ' Lastest Cal Data from the MfcCalibrationData table - the Normal way
Global Const calHistorical = 2              ' Historical Cal Data from the MfcCalibrationData table
Global Const calTestCurr = 8                ' Cal Data from the zCurrCalibration table - for Testing during calibration
' Constants to indicate calibration mode operation
Global Const CalOpeningValves = 1
Global Const CalNormalOperation = 2
Global Const CalClosingValves = 3
' New-Calibration-Sequence constants
Global Const newcalIdle = 0
Global Const newcalCalcheck_AsFound = 1
Global Const newcalCalibrate = 2
Global Const newcalCalcheck_AsLeft = 3
Global Const newcalPrintReport = 4
Global Const newcalDisplay = 5
' Calibration method constants
Global Const calmetUndefined = 0            ' no specified calibration method
Global Const calmetRawOnly = 1              ' Enter Raw Values Only; why??
Global Const calmetActualOnly = 2           ' Enter Actual Values Only; the Normal way for MFC's and Scales
Global Const calmetRawAndActual = 3         ' Device is calibrated on a bench; the Normal way for Transducers, TCs & RTDs

' STATION MODES
Global Const MAX_MODE = 39                  ' max mode number; for array sizing
Global Const VBIDLE = 1                     ' Idle
Global Const VBIDLEWAITING = 2              ' Waiting to Idle
Global Const VBCOMPLETE = 3                 ' Test is complete, generating reports, etc.
Global Const VBSCALEWAIT = 4                ' Wait for an available scale
Global Const VBSHIFTWAIT = 5                ' Shift wait
Global Const VBSTARTWAIT = 6                ' Delayed Start
Global Const VBLEAKWAIT = 7                 ' Leak check pause for LCP available transducer
Global Const VBPURGEWAIT = 8                ' Waiting to Purge
Global Const VBPAUSEVACSW = 9               ' Paused for System Vacuum Switch Off
Global Const VBLEAK = 10                    ' LeakCheck
Global Const VBPOSTLEAK = 11                ' Post LeakCheck Pause
Global Const VBWBPAUSE = 12                 ' Load pause for WaterBath Temp
Global Const VBPRELOAD = 13                 ' Preparing for Load  - N2 purge of load lines
Global Const VBLOAD = 14                    ' Load
Global Const VBPOSTLOAD = 15                ' Post Load Pause
Global Const VBPOSTLOADOPER = 16            ' Post Load Pause for Operator
Global Const VBPURGE = 17                   ' Purge
Global Const VBPOSTPURGE = 18               ' Post Purge Pause
Global Const VBPOSTPURGEOPER = 19           ' Post Purge Pause for Operator
Global Const VBPURGECONT = 20               ' Continue Purge (from alarm or oot; waiting for PurgeAir Ready)
Global Const VBCOURSEWAIT = 21              ' Wait until Operator gives OK (JobSequence Course = WaitForOK)
Global Const VBCOURSEPAUSE = 22             ' Pause (JobSequence Course = Pause)
Global Const VBPAUSE = 23                   ' Pause (by Recipe)
Global Const VBPAUSEBYUSER = 24             ' Pause initiated by User pressing the "Pause" pushbutton
Global Const VBGASPAUSE = 25                ' Load pause for gas replacement
Global Const VBFIDPAUSE = 26                ' Load pause for FID
Global Const VBLEAKERROR = 27               ' Leak Check Error
Global Const VBLEAKTEST = 31                ' LeakTest
Global Const VBPAUSEOOT = 38                ' Paused for OOT condition
Global Const VBPAUSEALARM = 39              ' Pause for General Alarm

' AUTO/MAN CONSTANTS
Global Const VBAUTO = 0                     ' Automatic
Global Const VBMANUAL = 1                   ' Manual

' CYCLE TYPE CONSTANTS
Global Const CycleUndefined = 0
Global Const CyclePurgeLoad = 1
Global Const CycleLoadPurge = 2

' LEAKCHECK PHASE CONSTANTS
Global Const LeakPurging = 0
Global Const LeakPressurizing = 1
Global Const LeakTesting = 2
Global Const LeakComplete = 3
          
' LOAD PHASE CONSTANTS
Global Const LoadStarting = 0
Global Const LoadLoading = 1
Global Const LoadComplete = 2
Global Const LoadStopping = 3
Global Const LoadPause = 7
Global Const LoadPrep = 9
          
' PURGE PHASE CONSTANTS
Global Const PurgeStarting = 0
Global Const PurgePurging = 1
Global Const PurgeComplete = 2
Global Const PurgeStopping = 3
Global Const PurgePause = 7
          
' SEQUENCE CONSTANTS
Global Const seqIdle = 0
Global Const seqCanVentN2Feed = 1
Global Const seqLeakTest = 4

' Reporting Constants                         Needed for reporting
Global Const FILEPAGECOLS = 100             ' Number of columns in file page
Global Const FILEPAGELINES = 66             ' Number of lines in file page
Global Const RTMARGIN = 3
Global Const LTMARGIN = 3
Global Const TOPMARGIN = 5
Global Const BOTTOMMARGIN = 5               ' Footer is printed after bottom margin
Global Const SUMMARYREPORT = 1              ' Summary Report
Global Const DETAILREPORT = 2               ' Detail report

' Thermocouple constants
Global Const TCTypeB = 0                    ' Range = +42C  to +1820
Global Const TCTypeJ = 1                    ' Range = -270C to +1200
Global Const TCTypeK = 2                    ' Range = -270C to +1372
Global Const TCTypeR = 3                    ' Range = -50C  to +1768
Global Const TCTypeS = 4                    ' Range = -50C  to +1768
Global Const TCTypeT = 5                    ' Range = -270C to +400

' DelayBox constants
Global Const msgSHOW = True                 ' show Screen
Global Const msgNOSHOW = False              ' don't show, just load

' System Timer constants
Global Const tmrScanIO = 1                  ' Scan I/O Timer
Global Const tmrScales = 2                  ' Read Scales Timer
Global Const tmrAlmOOT = 3                  ' Alarm/OOT Timer
Global Const tmrDataLog = 4                 ' Data Logger Timer
Global Const tmrControl = 5                 ' Controllers Logic Timer
Global Const tmrStnLogic = 6                ' Stations Logic Timer
Global Const tmrSysTmr = 7                  ' System Timers Timer
Global Const tmrUnused8 = 8                 ' unused Timer
Global Const tmrUnused9 = 9                 ' unused Timer

' AO Output control constants
Global Const outZERO = 0                    ' set Output to 0%
Global Const outNORMAL = 1                  ' Normal Output

' Waterbath Chiller control constants
Global Const pumpZERO = 0                    ' set pump speed to 0%
Global Const pumpSLOW = 2                    ' High pump speed
Global Const pumpNORMAL = 4                  ' Normal pump speed
Global Const pumpFAST = 6                    ' High pump speed

' Controller Index Constants
Public Enum PidControlIndex
    pasTEMPERATURE = 1                      ' PAS Temperature Local Control
    pasMOISTURE = 2                         ' PAS Moisture Local Control
    wbSuperTemp = 5                         ' WaterBath Supervisory Temperature Control
    stn1LoadRate = 11                       ' Stn #1 LoadRate Control
    stn2LoadRate = 12                       ' Stn #2 LoadRate Control
    stn3LoadRate = 13                       ' Stn #3 LoadRate Control
    stn4LoadRate = 14                       ' Stn #4 LoadRate Control
    stn5LoadRate = 15                       ' Stn #5 LoadRate Control
    stn6LoadRate = 16                       ' Stn #6 LoadRate Control
    stn7LoadRate = 17                       ' Stn #7 LoadRate Control
    stn8LoadRate = 18                       ' Stn #8 LoadRate Control
    stn9LoadRate = 19                       ' Stn #9 LoadRate Control
    stn1LeakTest = 21                       ' Stn #1 LeakTest Control
    stn2LeakTest = 22                       ' Stn #2 LeakTest Control
    stn3LeakTest = 23                       ' Stn #3 LeakTest Control
    stn4LeakTest = 24                       ' Stn #4 LeakTest Control
    stn5LeakTest = 25                       ' Stn #5 LeakTest Control
    stn6LeakTest = 26                       ' Stn #6 LeakTest Control
    stn7LeakTest = 27                       ' Stn #7 LeakTest Control
    stn8LeakTest = 28                       ' Stn #8 LeakTest Control
    stn9LeakTest = 29                       ' Stn #9 LeakTest Control
End Enum

' MsComm Index constants
Global Const mscommChiller = 1              ' Chiller MsComm index

' Text constants
Global Const BLANK = ""

' Value constants
Global Const VALUE0 = 0
Global Const VALUE1 = 1
Global Const VALUE2 = 2
Global Const VALUE3 = 3
Global Const VALUE4 = 4
Global Const VALUE5 = 5
Global Const VALUE6 = 6
Global Const VALUE7 = 7
Global Const VALUE8 = 8
Global Const VALUE9 = 9
Global Const VALUE10 = 10

' System paused constants
Global Const NOTPAUSED = 0                  ' System Not Paused
Global Const SYSTEMPAUSED = 1               ' System Paused

' Configuration constants
'   AutoLogon Constants
Global Const autologonOFF = 0               ' Auto Logon is OFF
Global Const autologonON = 1                ' Auto Logon is ON
Global Const autologonAPS = 2               ' Auto Logon as APS

'   Response to Leak Check Failure
Global Const MANUALSTOP = 0                 ' STOP pushbutton is only choice (default)
Global Const MANUALCHOOSE = 1               ' Operator chooses either STOP or CONTINUE
Global Const AUTOSTOP = 2                   ' Automatically Stop the test
Global Const AUTOCONTINUE = 3               ' Automatically Continue the test

' Recipe start method constants
Global Const STARTNOW = 0                   ' No delay
Global Const STARTDELAYED = 1               ' Delay Start by x minutes
Global Const STARTATDATE = 2                ' Start at specified DateTime

' Recipe end method constants
Global Const ENDCYCLES = 0                  ' End after x cycles
Global Const ENDWEIGHTCHG = 1               ' End after Load Weight Change per Cycle stabilizes

' Recipe leakcheck method constants
Global Const NOLEAKCHECK = 0
Global Const LEAKCHECKPRI = 1
Global Const LEAKCHECKAUX = 2
          
' Recipe load method constants
Global Const NOLOAD = 0                     ' No load
Global Const LOADBYTIME = 1                 ' Load by Time
Global Const LOADBYWC = 2                   ' Load by Working Capacity
Global Const LOADBYWEIGHT = 3               ' Load by Weight
Global Const LOADBYBREAKTHRU = 4            ' Load by Breakthrough
Global Const LOADBYFID = 5                  ' Load by FID Breakthrough

' Recipe purge method constants
Global Const NOPURGE = 0                    ' No purge
Global Const PURGEBYTIME = 1                ' Purge by Time
Global Const PURGEBYVOLUME = 2              ' Purge by Canister Volumes
Global Const PURGEAUXONLY = 3               ' Purge (by Time) Aux Canister Only
Global Const PURGEBYPROFILE = 4             ' Purge by Flow/Time Profile
Global Const PURGEBYWC = 5                  ' Purge by (% of) Working Capacity
Global Const PURGETOTARGET = 6              ' Purge to a Target Weight
Global Const PURGETOUNDOLOAD = 7            ' Purge to Undo Load
Global Const PURGEBYLITERS = 8              ' Purge by Liters of PurgeAir Flow

' Recipe purge to target constants
Global Const NOTARGET = 0                   ' No purge-to-target
Global Const TARGETCONTINUOUS = 1           ' Continuous purge-to-target
Global Const TARGETPURGEPAUSE = 2           ' Purge/Pause/Repeat purge-to-target

' Recipe purge profile step type constants
Global Const NOSTEP = 0                     ' Undefined profile step
Global Const STEPSTEP = 1                   ' Step Mfc SP at end of step
Global Const STEPRAMP = 2                   ' Ramp Mfc SP during step
Global Const STEPLAST = 3                   ' This step is the Last Step (always 0 duration)

' Recipe/Canister constants
Global Const MASTERMODE = 0                 ' Canister/Recipe screen is editing Master Canisters/Recipes
Global Const STATIONMODE = 1                ' Canister/Recipe screen is editing Station Canisters/Recipes

' Job Sequence Course Type constants
Global Const courseUndefined = 0            ' undefined
Global Const courseWait = 1                 ' Wait for Operator
Global Const coursePause = 2                ' Pause
Global Const courseRecipe = 3               ' Run Recipe

' Selected PurgeProfile Destination constants
Global Const profdestUndefined = 0          ' undefined
Global Const profdestProfile = 1            ' PurgeProfile
Global Const profdestRecipe = 2             ' Recipe

' Selected Recipe Destination constants
Global Const rcpdestUndefined = 0           ' undefined
Global Const rcpdestCourse = 1              ' Job Sequence Course
Global Const rcpdestRecipe = 2              ' Recipe

' Test Stop(Abort) Codes
Global Const AUTO_STOP = 1                   ' Test is stopped by program logic
Global Const EXIT_STOP = 2                   ' Operator Exit of application
Global Const OPER_STOP = 3                   ' Operator pressed STOP button on Station Detail screen

'   Report Codes
Global Const NORMALUPDATE = 0               ' Normal write
Global Const LOADBEGIN = 1                  ' Beginning of Load
Global Const LOADDONE = 2                   ' End of Load
Global Const PURGEBEGIN = 3                 ' Beginning of Purge
Global Const PURGEDONE = 4                  ' End of Purge
Global Const LT_BEGIN = 8                   ' Begin of LeakTest
Global Const LT_DONE = 9                    ' End of LeakTest
Global Const OOTPAUSEBEGIN = 11             ' OOT Pause begins
Global Const OOTPAUSECLEAR = 12             ' OOT Pause cleared by operator
Global Const LCBEGINPHASE0 = 13             ' Leak Check phase 0 (purging) begins
Global Const LCBEGINPHASE1 = 14             ' Leak Check phase 1 (pressurizing) begins
Global Const LCBEGINPHASE2 = 15             ' Leak Check phase 2 (testing) begins
Global Const LCTESTRESULT = 16              ' Leak Check Result (pass, fail, aborted)
Global Const LCOPERCONTINUE = 18            ' Operator pressed CONTINUE after a Leak Check failed
Global Const LCAUTOCONTINUE = 19            ' Automatic CONTINUE after a Leak Check failed

'   Report Filename Codes
Global Const RPT_NOTHING = 0                ' nothing; (actually "_" )
Global Const RPT_JOBNUMBER = 1              ' Job Number; JobXXXXXX_
Global Const RPT_STARTDTS = 2               ' Job Start DateTime; YYYYMMDD_HHMMSS_
Global Const RPT_STNSHIFT = 3               ' Station & Shift; StationX_ShiftX_
Global Const RPT_OPERENTER = 4              ' Operator Entered; *_
Global Const RPT_REMTASKID = 5              ' Remote Task Order ID; *_

' PurgeAirGenerator Local Control Types
Global Const pagNone = 0                    ' none
Global Const pagAlone = 1                   ' Stand-Alone
Global Const pagMaster = 2                  ' AK Master
Global Const pagClient = 3                  ' AK Client

'   Leak Check Results
Global Const NORESULT = 0                   ' no leak check test result; result reporting not applicable
Global Const RESULTFAIL_PURGETIMEOUT = 1    ' leak check failed - Purge Timeout
Global Const RESULTFAIL_PRESSURETIMEOUT = 2 ' leak check failed - Pressurize Timeout
Global Const RESULTFAIL_LEAKRATE = 3        ' leak check failed - Excessive Leak Rate
Global Const RESULTABORTAUTO = 7            ' leak check aborted automatically
Global Const RESULTABORTOPER = 8            ' leak check aborted by operator
Global Const RESULTGOOD = 9                 ' leak check passed

' CHILLER COMMAND INDEX CONSTANTS
Global Const chillerIn_PV = 110
Global Const chillerIn_SP = 120
Global Const chillerIn_Out = 121
Global Const chillerIn_OperMode = 122
Global Const chillerIn_OvrTmpSp = 123
Global Const chillerIn_P = 140
Global Const chillerIn_I = 141
Global Const chillerIn_Mode = 160
Global Const chillerIn_Type = 170
Global Const chillerIn_Version = 171
Global Const chillerIn_Status = 190
Global Const chillerIn_Stat = 191
Global Const chillerOut_SP = 220
Global Const chillerOut_Out = 221
Global Const chillerOut_OperMode = 222
Global Const chillerOut_P = 240
Global Const chillerOut_I = 241
Global Const chillerOut_Mode = 260
Global Const chillerOut_Start = 281
Global Const chillerOut_Stop = 280

' color constants from Color Table
' ********************************
Global Const AliceBlue = &HFFF8F0
Global Const AntiqueWhite = &HD7EBFA
Global Const Aqua = &HFFFF00
Global Const Aquamarine = &HD4FF7F
Global Const Azure = &HFFFFF0
Global Const Beige = &HDCF5F5
Global Const Bisque = &HC4E4FF
Global Const Black = &H0&
Global Const BlanchedAlmond = &HCDEBFF
Global Const Blue = &HFF0000
Global Const BlueViolet = &HE22B8A
Global Const Brown = &H2A2AA5
Global Const BurlyWood = &H87B8DE
Global Const CadetBlue = &HA09E5F
Global Const Chartreuse = &HFF7F&
Global Const Chocolate = &H1E69D2
Global Const Coral = &H507FFF
Global Const CornFlowerBlue = &HED9564
Global Const CornSilk = &HDCF8FF
Global Const Crimson = &H3C14DC
Global Const Cyan = &HFFFF00
Global Const DarkBlue = &H8B0000
Global Const DarkCyan = &H8B8B00
Global Const DarkGoldenRod = &HB86B8
Global Const DarkGray = &HA9A9A9
Global Const DarkGreen = &H6400&
Global Const DarkKhaki = &H6BB7BD
Global Const DarkMagenta = &H8B008B
Global Const DarkOliveGreen = &H2F6B55
Global Const DarkOrange = &H8CFF&
Global Const DarkOrchid = &HCC3299
Global Const DarkRed = &H8B&
Global Const DarkSalmon = &H7A96E9
Global Const DarkSeaGreen = &H8BBC8F
Global Const DarkSlateBlue = &H8B3D48
Global Const DarkSlateGray = &H4F4F2F
Global Const DarkTurquoise = &HD1CE00
Global Const DarkViolet = &HD30094
Global Const DeepPink = &H9314FF
Global Const DeepSkyBlue = &HFFBF00
Global Const DimGray = &H696969
Global Const DodgerBlue = &HFF901E
Global Const FireBrick = &H2222B2
Global Const FloralWhite = &HF0FAFF
Global Const ForestGreen = &H228B22
Global Const Fuchsia = &HFF00FF
Global Const Gainsboro = &HDCDCDC
Global Const GhostWhite = &HFFF8F8
Global Const Gold = &HD7FF&
Global Const Goldenrod = &H20A5DA
Global Const Gray = &H808080
Global Const Green = &H8000&
Global Const GreenYellow = &H2FFFAD
Global Const Honeydew = &HF0FFF0
Global Const HotPink = &HB469FF
Global Const IndianRed = &H5C5CCD
Global Const Indigo = &HB2004B
Global Const Ivory = &HF0FFFF
Global Const Khaki = &H8CE6F0
Global Const Lavender = &HFAE6E6
Global Const LavenderBlush = &HF5F0FF
Global Const LawnGreen = &HFC7C&
Global Const LemonChiffon = &HCDFAFF
Global Const LightBlue = &HE6D8AD
Global Const LightCoral = &H8080F0
Global Const LightCyan = &HFFFFE0
Global Const LightGoldenrodYellow = &HD2FAFA
Global Const LightGreen = &H90EE90
Global Const LightGray = &HD3D3D3
Global Const LightPink = &HC1B6FF
Global Const LightSalmon = &H7AA0FF
Global Const LightSeaGreen = &HAAB220
Global Const LightSkyBlue = &HFACE87
Global Const LightSlateGray = &H998877
Global Const LightSteelBlue = &HDEC4B0
Global Const LightYellow = &HE0FFFF
Global Const Lime = &HFF00&
Global Const LimeGreen = &H32CD32
Global Const Linen = &HE6F0FA
Global Const Magenta = &HFF00FF
Global Const Maroon = &H80&
Global Const MediumAquamarine = &HAACD66
Global Const MediumBlue = &HCD0000
Global Const MediumOrchid = &HD355BA
Global Const MediumPurple = &HDB7093
Global Const MediumSeaGreen = &H71B33C
Global Const MediumSlateBlue = &HEE687B
Global Const MediumSpringGreen = &H9AFA00
Global Const MediumTurquoise = &HCCD148
Global Const MediumVioletRed = &H8515C7
Global Const MidnightBlue = &H701919
Global Const MintCream = &HFAFFF5
Global Const MistyRose = &HE1E4FF
Global Const Moccasin = &HB5E4FF
Global Const NavajoWhite = &HADDEFF
Global Const Navy = &H800000
Global Const Olive = &H8080&
Global Const Olivedrab = &H238E6B
Global Const Orange = &HA5FF&
Global Const OrangeRed = &H45FF&
Global Const Orchid = &HD670DA
Global Const PaleGoldenrod = &HAAE8EE
'Global Const PaleGreen = &H98FB98
Global Const PaleTurquoise = &HEEEEAF
Global Const PaleVioletRed = &H9370DB
Global Const PapayaWhip = &HD5EFFF
Global Const PeachPuff = &HB9DAFF
Global Const Peru = &H3F85CD
Global Const Pink = &HCBC0FF
Global Const Plum = &HDDA0DD
Global Const PowderBlue = &HE6E0B0
Global Const Purple = &H800080
Global Const Red = &HFF&
Global Const RosyBrown = &H8F8FBC
Global Const RoyalBlue = &HE16941
Global Const SaddleBrown = &H13458B
Global Const Salmon = &H7280FA
Global Const SandyBrown = &H60A4F4
Global Const SeaGreen = &H578B2E
Global Const Seashell = &HEEF5FF
Global Const Sienna = &H2D52A0
Global Const Silver = &HC0C0C0
Global Const SkyBlue = &HEBCE87
Global Const SlateBlue = &HCD5A6A
Global Const SlateGray = &H908070
Global Const Snow = &HFAFAFF
Global Const SpringGreen = &H7FFF00
Global Const SteelBlue = &HB48246
Global Const Tan = &H8CB4D2
Global Const Teal = &H808000
Global Const Thistle = &HD8BFD8
Global Const Tomato = &H4763FF
Global Const Turquoise = &HD0E040
Global Const Violet = &HEE82EE
Global Const Wheat = &HB3DEF5
Global Const White = &HFFFFFF
Global Const WhiteSmoke = &HF5F5F5
Global Const Yellow = &HFFFF&
Global Const YellowGreen = &H32CD9A

' ********************************
' color Constants from Color Pallet
Global Const SOFTBLUE = &HFF6830
Global Const SOFTBLUETOO = &HFF6000
' reds
Global Const PALERED = &HC0C0FF
Global Const LTRED = &H8080FF
Global Const MEDRED = &HFF&
Global Const DKRED = &HC0&
Global Const DK2RED = &H80&
Global Const DK3RED = &H40&
' oranges
Global Const PALEORANGE = &HC0E0FF
Global Const LTORANGE = &H80C0FF
Global Const MEDORANGE = &H80FF&
Global Const DKORANGE = &H40C0&
Global Const DK2ORANGE = &H4080&
Global Const DK3ORANGE = &H404080
' yellows
Global Const PALEYELLOW = &HC0FFFF
Global Const LTYELLOW = &H80FFFF
Global Const MEDYELLOW = &HFFFF&
Global Const DKYELLOW = &HC0C0&
Global Const DK2YELLOW = &H8080&
Global Const DK3YELLOW = &H4040&
' greens
Global Const PALEGREEN = &HC0FFC0
Global Const LTGREEN = &H80FF80
Global Const MEDGREEN = &HFF00&
Global Const DKGREEN = &HC000&
Global Const DK2GREEN = &H8000&
Global Const DK3GREEN = &H4000&
' cyans
Global Const PALECYAN = &HFFFFC0
Global Const LTCYAN = &HFFFF80
Global Const MEDCYAN = &HFFFF00
Global Const DKCYAN = &HC0C000
Global Const DK2CYAN = &H808000
Global Const DK3CYAN = &H404000
' blues
Global Const PALEBLUE = &HFFC0C0
Global Const LTBLUE = &HFF8080
Global Const MEDBLUE = &HFF0000
Global Const DKBLUE = &HC00000
Global Const DK2BLUE = &H800000
Global Const DK3BLUE = &H400000
' purples
Global Const PALEVIOLET = &HFFC0FF
Global Const LTPURPLE = &HE0E0E0
Global Const MEDPURPLE = &HFF00FF
Global Const DKPURPLE = &HC000C0
Global Const DK2PURPLE = &H800080
Global Const DK3PURPLE = &H400040
' blacks, grays, whites
Global Const PALEGRAY = &HFFFFFF    ' aka white
Global Const LTGRAY = &HE0E0E0
Global Const MEDGRAY = &HC0C0C0
Global Const DKGRAY = &H808080
Global Const DK2GRAY = &H404040
Global Const DK3GRAY = &H0&         ' aka black

' // index definition for OOT Response Codes
Public Enum OotRespCodeIndex
    ootrspUndefined = 0                 ' //    undefined
    ootrspContinue                      ' //    continue
    ootrspPause                         ' //    pause
    ootrspStop                          ' //    stop
End Enum

' // index definition for LiveFuel State Codes
Public Enum LiveFuelStateIndex
    fuelDead = 0                        ' //    Dead; replace immediately
    fuelWeak                            ' //    Weak; replace before next load
    fuelOK                              ' //    OK
End Enum

' // index definition for WaterBath Control Codes
Public Enum WaterBathControlIndex
    wbDirect = 0                        ' //    Direct setting of WB SetPoint
    wbFuelTemp = 1                     ' //    Use WaterBath to control Live Fuel Tank Temp
    wbVaporTemp = 2                    ' //    Use WaterBath to control Live Fuel Vapor Temp
End Enum

' OOT Count indexes
Global Const ootNone = 0
Global Const ootBtnFlow = 1
Global Const ootNitFlow = 2
Global Const ootFuelTemp = 3
Global Const ootPurFlow = 4
Global Const ootAirMoist = 5
Global Const ootAirTemp = 6
Global Const ootCanVent = 7
Global Const ootLoadRate = 8
Global Const ootPurgeDp = 9
Global Const ootFuelLevel = 10
Global Const ootStorageLevel = 11
Global Const ootPurgeOvenTemp = 12
Global Const ootWaterBathTemp = 13

' Mass Flow Controller constants
Global Const MFCBUTANE = 0
Global Const MFCNITROGEN = 1
Global Const MFCPURGEAIR = 2
Global Const MFCORVRBUT = 3
Global Const MFCORVRNIT = 4
Global Const MFCORVRPRG = 5
Global Const MFCLIVEFUEL = 6
Global Const MFCORVRLIVE = 7
Global Const MAXMFC = 7                             ' Max Allowed value for Mass Flow Controller Index

' Station Type constants
Global Const STN_REGULAR_TYPE = 1                   ' Regular Station Type
Global Const STN_ORVR_TYPE = 2                      ' ORVR Station Type
Global Const STN_LIVEFUEL_TYPE = 3                  ' LiveFuel Station Type
Global Const STN_ORVR2_TYPE = 4                     ' ORVR with 2 Load MFC's (per gas except PurgeAir) Station Type
Global Const STN_LIVEREG_TYPE = 6                   ' Regular & LiveFuel Combo Station Type
Global Const STN_LIVEORVR2_TYPE = 7                 ' Regular & LiveFuel with 2 Load MFC's (per gas except PurgeAir) Combo Station Type
Global Const STN_COMBO3_TYPE = 8                    ' "future" Combo Station Type
Global Const STN_DUMMY_TYPE = 9                     ' Dummy Station Type
Global Const STN_LEAKTEST_TYPE = 17                 ' Leak Test Station Type (40 CFR 1066.985; no loading, no purgeing)
Global Const MAX_STNTYPE = 19                       ' Max Station Type Index

' Opto interface constants
Global Const MAX_NODE = 11                          ' The highest number for the Opto Node Number
Global Const MAX_ADDR = 49                          ' The highest number for the Opto Address Number
Global Const MAX_CHAN = 15                          ' The highest number for the Opto Channel Number
Global Const MAX_SLOT = 15                          ' The highest number for the Opto Board Slot Number
Global Const MAX_ANA_COM = 29                       ' The highest number for the Common Analog IO Function Index
Global Const MAX_DIG_COM = 99                       ' The highest number for the Common Digital IO Function Index
Global Const MAX_ANA_FID = 19                       ' The highest number for the FID Analog IO Function Index
Global Const MAX_DIG_FID = 39                       ' The highest number for the FID Digital IO Function Index
Global Const MAX_ANA_PRG = 19                       ' The highest number for the PurgeAir Source Analog IO Function Index
Global Const MAX_DIG_PRG = 19                       ' The highest number for the PurgeAir Source Digital IO Function Index
Global Const MAX_ANA_STN = 39                       ' The highest number for the Station Analog IO Function Index
Global Const MAX_DIG_STN = 99                       ' The highest number for the Station Digital IO Function Index

' ANALOGS
' ANALOGS
' ANALOGS

' COMMON ANALOG IO FUNCTION INDEX CONSTANTS
Global Const acPasTempSensor = 1                    ' Index into the Common Analog IO Functions
Global Const acPasHumiditySensor = 2                ' Index into the Common Analog IO Functions
Global Const acAmbTempSensor = 5                    ' Index into the Common Analog IO Functions
Global Const acAmbHumiditySensor = 6                ' Index into the Common Analog IO Functions
Global Const acComnPressSensor = 7                  ' Index into the Common Analog IO Functions
Global Const acAmbBaroSensor = 8                    ' Index into the Common Analog IO Functions
Global Const acCustCalDevice = 10                   ' Index into the Common Analog IO Functions
Global Const acCommonTC1 = 11                       ' Index into the Common Analog IO Functions
Global Const acCommonTC2 = 12                       ' Index into the Common Analog IO Functions
Global Const acCommonTC3 = 13                       ' Index into the Common Analog IO Functions
Global Const acCommonTC4 = 14                       ' Index into the Common Analog IO Functions
Global Const acCommonTC5 = 15                       ' Index into the Common Analog IO Functions
Global Const acCommonTC6 = 16                       ' Index into the Common Analog IO Functions
Global Const acPASMoistCntrlOut = 20                ' Index into the Common Analog IO Functions

' PURGEAIR SOURCE ANALOG IO FUNCTION INDEX CONSTANTS
Global Const apTemp = 1                             ' Index into the PurgeAir Source Analog IO Functions
Global Const apHumidity = 2                         ' Index into the PurgeAir Source Analog IO Functions
Global Const apBaro = 3                             ' Index into the PurgeAir Source Analog IO Functions

' STATION ANALOG IO FUNCTION INDEX CONSTANTS
Global Const asNitrogenFlowSP = 1                   ' Index into the Station Analog IO Functions
Global Const asButaneFlowSP = 2                     ' Index into the Station Analog IO Functions
Global Const asPurgeAirFlowSP = 3                   ' Index into the Station Analog IO Functions
Global Const asNitrogenORVRFlowSP = 4               ' Index into the Station Analog IO Functions
Global Const asButaneORVRFlowSP = 5                 ' Index into the Station Analog IO Functions
Global Const asPurgeAirORVRFlowSP = 6               ' Index into the Station Analog IO Functions
Global Const asLiveFuelVaporFlowSP = 7              ' Index into the Station Analog IO Functions
Global Const asLiveFuelVaporORVRFlowSP = 8          ' Index into the Station Analog IO Functions
Global Const asPurgeDiffPress = 9                   ' Index into the Station Analog IO Functions
Global Const asNitrogenFlow = 11                    ' Index into the Station Analog IO Functions
Global Const asButaneFlow = 12                      ' Index into the Station Analog IO Functions
Global Const asPurgeAirFlow = 13                    ' Index into the Station Analog IO Functions
Global Const asNitrogenORVRFlow = 14                ' Index into the Station Analog IO Functions
Global Const asButaneORVRFlow = 15                  ' Index into the Station Analog IO Functions
Global Const asPurgeAirORVRFlow = 16                ' Index into the Station Analog IO Functions
Global Const asLiveFuelVaporFlow = 17               ' Index into the Station Analog IO Functions
Global Const asLiveFuelVaporORVRFlow = 18           ' Index into the Station Analog IO Functions
Global Const asLoadPressure = 19                    ' Index into the Station Analog IO Functions
Global Const asPurgeOvenTempSP = 21                 ' Index into the Station Analog IO Functions
Global Const asPurgeOvenTemp = 22                   ' Index into the Station Analog IO Functions
Global Const asFuelHeaterSP = 24                    ' Index into the Station Analog IO Functions
Global Const asFuelHeaterTemp = 25                  ' Index into the Station Analog IO Functions
Global Const asFuelTankTemp = 26                    ' Index into the Station Analog IO Functions
Global Const asFuelTankLevel = 27                   ' Index into the Station Analog IO Functions
Global Const asFuelVaporTemp = 28                   ' Index into the Station Analog IO Functions
Global Const asStorageTankLevel = 29                ' Index into the Station Analog IO Functions
Global Const asStationTC1 = 31                      ' Index into the Station Analog IO Functions
Global Const asStationTC2 = 32                      ' Index into the Station Analog IO Functions
Global Const asLtInletPress = 36                    ' Index into the Station Analog IO Functions
Global Const asLtN2Temp = 37                        ' Index into the Station Analog IO Functions

' DIGITALS
' DIGITALS
' DIGITALS

' COMMON DIGITAL IO FUNCTION INDEX CONSTANTS
Global Const icHornSilencePB = 1                    ' Index into the Common Digital IO Functions
Global Const icExhaustFlowFS = 2                    ' Index into the Common Digital IO Functions
Global Const icEStopSw = 3                          ' Index into the Common Digital IO Functions
Global Const icMaintSw = 4                          ' Index into the Common Digital IO Functions
Global Const icDoorSw = 5                           ' Index into the Common Digital IO Functions
Global Const ic20LelGasSw = 7                       ' Index into the Common Digital IO Functions
Global Const icUpsFaultSw = 8                       ' Index into the Common Digital IO Functions
Global Const icUpsActiveSw = 9                      ' Index into the Common Digital IO Functions
Global Const icSystemVacSw = 10                     ' Index into the Common Digital IO Functions
Global Const icAlarmBeacon = 11                     ' Index into the Common Digital IO Functions
Global Const icAlarmHorn = 12                       ' Index into the Common Digital IO Functions
Global Const icPauseLT = 13                         ' Index into the Common Digital IO Functions
Global Const icButaneShutoffSol = 15                ' Index into the Common Digital IO Functions
Global Const icLeakCheckExhaustSol = 16             ' Index into the Common Digital IO Functions
Global Const icExtAlmContactSw = 17                 ' Index into the Common Digital IO Functions
Global Const icCustLowGasSw = 18                    ' Index into the Common Digital IO Functions
Global Const icPurgeRequestOut = 20                 ' Index into the Common Digital IO Functions
Global Const icPurgeReadyIn = 21                    ' Index into the Common Digital IO Functions
Global Const icPASRequestIn = 30                    ' Index into the Common Digital IO Functions
Global Const icPASReadyOut = 31                     ' Index into the Common Digital IO Functions
Global Const icRunLocalPASIn = 33                   ' Index into the Common Digital IO Functions
Global Const icPASHeaterSSR = 34                    ' Index into the Common Digital IO Functions
Global Const icPurgeDryAirSupplySol = 38            ' Index into the Common Digital IO Functions
Global Const icPurgeAirSourceSelectSol = 39         ' Index into the Common Digital IO Functions
Global Const icScale01AuxAirSol = 51                ' Index into the Common Digital IO Functions
Global Const icScale02AuxAirSol = 52                ' Index into the Common Digital IO Functions
Global Const icScale03AuxAirSol = 53                ' Index into the Common Digital IO Functions
Global Const icScale04AuxAirSol = 54                ' Index into the Common Digital IO Functions
Global Const icScale05AuxAirSol = 55                ' Index into the Common Digital IO Functions
Global Const icScale06AuxAirSol = 56                ' Index into the Common Digital IO Functions
Global Const icScale07AuxAirSol = 57                ' Index into the Common Digital IO Functions
Global Const icScale08AuxAirSol = 58                ' Index into the Common Digital IO Functions
Global Const icScale09AuxAirSol = 59                ' Index into the Common Digital IO Functions
Global Const icScale10AuxAirSol = 60                ' Index into the Common Digital IO Functions
Global Const icScale11AuxAirSol = 61                ' Index into the Common Digital IO Functions
Global Const icScale12AuxAirSol = 62                ' Index into the Common Digital IO Functions
Global Const icScale13AuxAirSol = 63                ' Index into the Common Digital IO Functions
Global Const icScale14AuxAirSol = 64                ' Index into the Common Digital IO Functions
Global Const icScale15AuxAirSol = 65                ' Index into the Common Digital IO Functions
Global Const icScale16AuxAirSol = 66                ' Index into the Common Digital IO Functions
Global Const icLiveFuelPurgePS = 71                 ' Index into the Common Digital IO Functions

' PURGEAIR SOURCE DIGITAL IO FUNCTION INDEX CONSTANTS
Global Const ipPiabSol = 1                          ' Index into the PurgeAir Source Digital IO Functions
Global Const ipPurgeVacuumSw = 2                    ' Index into the PurgeAir Source Digital IO Functions
' Global Const ipPurgeRequestOut = 3                  ' Now in Common Digital IO Functions (max one per system)
' Global Const ipPurgeReadyIn = 4                     ' Now in Common  Digital IO Function (max one per system)
Global Const ipAuxAirSol = 5                        ' Index into the PurgeAir Source Digital IO Functions
Global Const ipPosPrsPrgSol = 6                     ' Index into the PurgeAir Source Digital IO Functions

' STATION DIGITAL IO FUNCTION INDEX CONSTANTS
Global Const isNitrogenSol = 1                      ' Index into the Station Digital IO Functions
Global Const isButaneSol = 2                        ' Index into the Station Digital IO Functions
Global Const isPurgeSol = 3                         ' Index into the Station Digital IO Functions
Global Const isPriDirectionSol = 4                  ' Index into the Station Digital IO Functions
Global Const isAuxCanVentSol = 5                    ' Index into the Station Digital IO Functions
Global Const isLeakCheckSol = 6                     ' Index into the Station Digital IO Functions
Global Const isAuxPurgeSol = 7                      ' Index into the Station Digital IO Functions
Global Const isPriAuxVentSol = 8                    ' Index into the Station Digital IO Functions
Global Const isAuxDirectionSol = 9                  ' Index into the Station Digital IO Functions
Global Const isAuxLeakCheckSol = 10                 ' Index into the Station Digital IO Functions
Global Const isPauseLT = 11                         ' Index into the Station Digital IO Functions
Global Const isIdleLT = 12                          ' Index into the Station Digital IO Functions
Global Const isPriSeriesPurgeSol = 13               ' Index into the Station Digital IO Functions
Global Const isAuxSeriesPurgeSol = 14               ' Index into the Station Digital IO Functions
Global Const isLoadShift2Sol = 16                   ' Index into the Station Digital IO Functions
Global Const isVentShift2Sol = 17                   ' Index into the Station Digital IO Functions
Global Const isPurgeShift2Sol = 18                  ' Index into the Station Digital IO Functions
Global Const isNitrogenOrvrSol = 21                 ' Index into the Station Digital IO Functions
Global Const isButaneOrvrSol = 22                   ' Index into the Station Digital IO Functions
' Global Const isPurgeOrvrSol = 23                    ' Now same as isPurgeSol Station Digital IO Function
Global Const isLiveFuelOrvrSol = 24                 ' Index into the Station Digital IO Functions
Global Const isLoadTypeSelectSol = 28               ' Index into the Station Digital IO Functions
Global Const isLiveFuelSol = 29                     ' Index into the Station Digital IO Functions
Global Const isFuelVentSol = 30                     ' Index into the Station Digital IO Functions
Global Const isFuelRecircSol = 31                   ' Index into the Station Digital IO Functions
Global Const isFuelDrainSol = 32                    ' Index into the Station Digital IO Functions
Global Const isFuelFillSol = 33                     ' Index into the Station Digital IO Functions
Global Const isFuelPressSol = 34                    ' Index into the Station Digital IO Functions
Global Const isFuelVaporSol = 35                    ' Index into the Station Digital IO Functions
Global Const isFuelPumpMotor = 36                   ' Index into the Station Digital IO Functions
Global Const isFuelHiHiLevelLS = 37                 ' Index into the Station Digital IO Functions
Global Const isFuelHighLevelLS = 38                 ' Index into the Station Digital IO Functions
Global Const isFuelLowLevelLS = 39                  ' Index into the Station Digital IO Functions
Global Const isFuelHeaterSSR = 40                   ' Index into the Station Digital IO Functions
Global Const isFuelOverTempSw = 41                  ' Index into the Station Digital IO Functions
Global Const isSheathOverTempSw = 42                ' Index into the Station Digital IO Functions
Global Const isFuelSafetyLevelLS = 43               ' Index into the Station Digital IO Functions
Global Const isFuelPressPS = 44                     ' Index into the Station Digital IO Functions
Global Const isStorageHiHiLevelLS = 45              ' Index into the Station Digital IO Functions
Global Const isStorageLowLevelLS = 46               ' Index into the Station Digital IO Functions
Global Const isStorageDrainSol = 47                 ' Index into the Station Digital IO Functions
Global Const isStorageFillSol = 48                  ' Index into the Station Digital IO Functions
Global Const isStorageFillRequest = 49              ' Index into the Station Digital IO Functions
Global Const isFuelOverTempResetOut = 50            ' Index into the Station Digital IO Functions
Global Const isCanVentAlarmSw = 51                  ' Index into the Station Digital IO Functions
Global Const isPurgeLocationSupplySelectSol = 58    ' Index into the Station Digital IO Functions
Global Const isPurgeLocationVentSelectSol = 59      ' Index into the Station Digital IO Functions
Global Const isLoadShift3Sol = 62                   ' Index into the Station Digital IO Functions
Global Const isVentShift3Sol = 63                   ' Index into the Station Digital IO Functions
Global Const isPurgeShift3Sol = 64                  ' Index into the Station Digital IO Functions
Global Const isLoadShift4Sol = 66                   ' Index into the Station Digital IO Functions
Global Const isVentShift4Sol = 67                   ' Index into the Station Digital IO Functions
Global Const isPurgeShift4Sol = 68                  ' Index into the Station Digital IO Functions
Global Const isAuxOutput1 = 71                      ' Index into the Station Digital IO Functions
Global Const isAuxOutput2 = 72                      ' Index into the Station Digital IO Functions
Global Const isAuxOutput3 = 73                      ' Index into the Station Digital IO Functions
Global Const isAuxOutput4 = 74                      ' Index into the Station Digital IO Functions


' Declare array for File Maintenance List
Type fList
  fName As String
  fDate As Date
End Type


' *************************************************************************************
' Declare LSQ Calibration Data Type
'
'
Type LsqCalData
    X As Single
    X2 As Single
    X3 As Single
    X4 As Single
    X5 As Single
    X6 As Single
    R2 As Single
End Type

' *************************************************************************************
' Declare (AI, MFC or Scale) Calibration Point Data Type
'
'
Type CalDataPoint
    ActualValue As Single
    RawValue As Single
    ActualPercent As Single
    RawPercent As Single
End Type

' *************************************************************************************
' Declare (AI, MFC or Scale) CalCheck Point Data Type
'
'
Type CalCheckDataPoint
    ActualValue As Single
    DesiredValue As Single
    ActualPercent As Single
    DesiredPercent As Single
    PercentDiff As Single
End Type

' *************************************************************************************
' Declare AI Calibration Type
'
'
Type AICalibration
    dts As Date
    StandardTempValue As Single
    StandardTempUnits As String
    StandardPressValue As Single
    StandardPressUnits As String
    CalibratedBy As String
    Equipment As String
    Comment As String
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalDataPoint
    CalData As LsqCalData
End Type
' Analog Input calibration variables
Global Com_AiCal(0 To MAX_ANA_COM) As AICalibration
Global Prg_AiCal(1 To MAX_PRG, 0 To MAX_ANA_PRG) As AICalibration
Global Stn_AiCal(1 To MAX_STN, 0 To MAX_ANA_STN) As AICalibration
Global PrevCom_AiCal(0 To MAX_ANA_COM) As AICalibration
Global PrevPrg_AiCal(1 To MAX_PRG, 0 To MAX_ANA_PRG) As AICalibration
Global PrevStn_AiCal(1 To MAX_STN, 0 To MAX_ANA_STN) As AICalibration

' *************************************************************************************
' Declare AI CalCheck Type
'
'
Type AICalCheck
    dts As Date
    calDTS As Date
    CalibratedBy As String
    Equipment As String
    Comment As String
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalCheckDataPoint
End Type
' Analog Input calcheck variables
Global Com_AiCalChecks(0 To MAX_ANA_COM, 0 To MAXCALCHECKS) As AICalCheck
Global Prg_AiCalChecks(1 To MAX_PRG, 0 To MAX_ANA_PRG, 0 To MAXCALCHECKS) As AICalCheck
Global Stn_AiCalChecks(1 To MAX_STN, 0 To MAX_ANA_STN, 0 To MAXCALCHECKS) As AICalCheck
Global Com_AiCalCheckIdx(0 To MAX_ANA_COM) As Integer
Global Prg_AiCalCheckIdx(1 To MAX_PRG, 0 To MAX_ANA_PRG) As Integer
Global Stn_AiCalCheckIdx(1 To MAX_STN, 0 To MAX_ANA_STN) As Integer

' *************************************************************************************
' Declare Scale Calibration Type
'
'
Type SclCalibration
    dts As Date
    StandardTempValue As Single
    StandardTempUnits As String
    StandardPressValue As Single
    StandardPressUnits As String
    CalibratedBy As String
    Equipment As String
    Comment As String
    CalRangeMax As Single
    CalRangeMin As Single
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalDataPoint
    CalData As LsqCalData
End Type
' Scale calibration variables
Global Scale_Cal(0 To MAX_SCALES) As SclCalibration
Global PrevScale_Cal(0 To MAX_SCALES) As SclCalibration

' *************************************************************************************
' Declare Scale CalCheck Type
'
'
Type SclCalCheck
    dts As Date
    calDTS As Date
    CalibratedBy As String
    Equipment As String
    Comment As String
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalCheckDataPoint
End Type
' Scale calcheck variables
Global Scale_CalChecks(0 To MAX_SCALES, 0 To MAXCALCHECKS) As SclCalCheck
Global Scale_CalCheckIdx(0 To MAX_SCALES) As Integer
' *************************************************************************************
' Declare MFC Calibration Type
'
'
Type MfcCalibration
    dts As Date
    StandardTempValue As Single
    StandardTempUnits As String
    StandardPressValue As Single
    StandardPressUnits As String
    CalibratedBy As String
    Equipment As String
    Comment As String
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalDataPoint
    CalData As LsqCalData
End Type
' mass flow controller calibration variables
Global Stn_MfcCal(1 To MAX_STN, 0 To MAXMFC) As MfcCalibration
Global PrevStn_MfcCal(1 To MAX_STN, 0 To MAXMFC) As MfcCalibration

' *************************************************************************************
' Declare MFC CalCheck Type
'
'
Type MfcCalCheck
    dts As Date
    calDTS As Date
    CalibratedBy As String
    Equipment As String
    Comment As String
    Method As Integer
    RawInputType As Integer
    NumPoints As Integer
    PointData(1 To MAXLSQCALPOINTS) As CalCheckDataPoint
End Type
' Scale calcheck variables
Global Stn_MfcCalChecks(1 To MAX_STN, 0 To MAXMFC, 0 To MAXCALCHECKS) As MfcCalCheck
Global Stn_MfcCalCheckIdx(1 To MAX_STN, 0 To MAXMFC) As Integer
' *******************************************************************************

' *******************************************************************************
' Declare arrays for Opto I/O Information
'
' *******************************************************************************
' Type 0     No Module
' Type 1     DI
' Type 2     DO
' Type 3     AI
' Type 4     AO
' Type 5     TC Type J
' Type 6     TC Type K
' Type 7     RTD 100 Ohms
' *******************************************************************************
Global Const optotypeNoModule = 0
Global Const optotypeDI = 1
Global Const optotypeDO = 2
Global Const optotypeAI = 3
Global Const optotypeAO = 4
Global Const optotypeTcJ = 5
Global Const optotypeTcK = 6
Global Const optotypeRTD = 7
' *******************************************************************************
' TypeCode 0     Not used
' TypeCode 5     Type J thermocouple
' TypeCode 7     0 to 10 volt dc (not preferred one)
' TypeCode 8     Type K thermocouple
' TypeCode 12    +/- 10 volt Blue module AI inputs
' TypeCode 133   0 to 10 volt dc (Preferred) Green module AO outputs
' TypeCode 100   SNAP Digital In
' TypeCode 180   SNAP Digital Out
' TypeCode 0A    RTD 100 ohm
' *******************************************************************************
Global Const optocodeUndefined = 0
Global Const optocodeTcJ = 5
Global Const optocode0to10vdc = 7
Global Const optocodeTcK = 8
Global Const optocodeAI = 12
Global Const optocodeAO = 133
Global Const optocodeSnapDI = 100
Global Const optocodeSnapDO = 180
Global Const optocodeRTD = 10
' *******************************************************************************
Type OptoAnalogHardware
  Type As Integer                                                   ' Channel Type; see above
  TypeCode As Integer                                               ' Opto Code; see above
  RawValue As Long                                                  ' Current Value as reported by Opto
End Type
Type OptoDigitalHardware
  Type As Integer                                                   ' Channel Type; see above
  TypeCode As Integer                                               ' Opto Code; see above
  RawValue As Boolean                                               ' Current Value as reported by Opto
End Type
Global OptoAIO(0 To MAX_ADDR, 0 To MAX_CHAN) As OptoAnalogHardware   ' Opto Analog I/O Information Array
Global OptoDIO(0 To MAX_ADDR, 0 To MAX_CHAN) As OptoDigitalHardware  ' Opto Digital I/O Information Array
Global OptoChanMask(0 To MAX_ADDR) As Long                           ' Opto Channel-In-Use Masks
Global OptoChanDesc(0 To MAX_ADDR, 0 To MAX_CHAN) As String          ' Opto Channel Functional Descriptor
Global OptoMaxNodeNum As Integer

' Declare arrays for Mapped I/O Information
'
Type MapAnalogIO
  desc As String                                                    ' Channel Description
  Type As Integer                                                   ' Channel Type; see above
  TypeCode As Integer                                               ' Opto Code; see above
  RawValue As Long                                                  ' Current Value as reported by Opto
  VdcMax As Single                                                  ' Vdc at EUMax
  VdcMin As Single                                                  ' Vdc at EUMin
  EuMax As Single                                                   ' Max Engr Unit Value
  EuMin As Single                                                   ' Min Engr Unit Value
  EUValue As Single                                                 ' Engr Unit Value after conversion
  CalData As LsqCalData
End Type
Type MapDigitalIO
  desc As String                                                    ' Channel Description
  Type As Integer                                                   ' Channel Type; see above
  TypeCode As Integer                                               ' Opto Code; see above
  RawValue As Boolean                                               ' Current Value as reported by Opto
  UseInverse As Boolean                                             ' True = Invert RawValue from Opto
  Value As Boolean                                                  ' Value (after inversion, if required)
End Type
Global Map_AIO(0 To MAX_ADDR, 0 To MAX_CHAN) As MapAnalogIO         ' Analog IO Information Array
Global Map_DIO(0 To MAX_ADDR, 0 To MAX_CHAN) As MapDigitalIO        ' Digital IO Information Array


' Declare arrays for Functional I/O Information
'
Type FuncAnalogIO
  addr As Integer                                                   ' Opto Address
  chan As Integer                                                   ' Opto Channel
  VdcMax As Single                                                  ' Vdc at EUMax
  VdcMin As Single                                                  ' Vdc at EUMin
  EuMax As Single                                                   ' Max Engr Unit Value
  EuMin As Single                                                   ' Min Engr Unit Value
  EUValue As Single                                                 ' Engr Unit Value after conversion
End Type
Type FuncDigitalIO
  addr As Integer                                                   ' Opto Address
  chan As Integer                                                   ' Opto Channel
  UseInverse As Boolean                                             ' True = Invert RawValue from Opto
  Value As Boolean                                                  ' Value (after inversion, if required)
End Type
Type ComFuncDefinition
  desc As String                                                    ' Description of this IO Function
  UsedIn As Boolean                                                 ' True = Used in this System
End Type
Type MfcFuncDefinition
  desc As String                                                    ' Description of this MFC type
  UsedIn(1 To MAX_STNTYPE) As Boolean                               ' True = Used in this Type of Station
End Type
Type StnFuncDefinition
  desc As String                                                    ' Description of this IO Function
  UsedIn(1 To MAX_STNTYPE) As Boolean                               ' True = Used in this Type of Station
End Type
Global WB_AIO As FuncAnalogIO                                                           ' WaterBath Analog IO Information Array
Global Com_AIO(0 To MAX_ANA_COM) As FuncAnalogIO                                        ' Common Analog IO Information Array
Global Com_DIO(0 To MAX_DIG_COM) As FuncDigitalIO                                       ' Common Digital IO Information Array
Global Prg_AIO(1 To MAX_PRG, 0 To MAX_ANA_PRG) As FuncAnalogIO                          ' PurgeAir Source Analog IO Information Array
Global Prg_DIO(1 To MAX_PRG, 0 To MAX_DIG_PRG) As FuncDigitalIO                         ' PurgeAir Source Digital IO Information Array
Global Stn_AIO(1 To MAX_STN, 0 To MAX_ANA_STN) As FuncAnalogIO                          ' Station Analog IO Information Array
Global Stn_DIO(1 To MAX_STN, 0 To MAX_DIG_STN) As FuncDigitalIO                         ' Station Digital IO Information Array
Global Com_AnaDef(0 To MAX_ANA_COM) As ComFuncDefinition                                ' Common Analog IO Function Definitions
Global Prg_AnaDef(0 To MAX_ANA_PRG) As ComFuncDefinition                                ' PurgeAir Source Analog IO Function Definitions
Global Stn_AnaDef(0 To MAX_ANA_STN) As StnFuncDefinition                                ' Station Analog IO Function Definitions
Global Com_DigDef(0 To MAX_DIG_COM) As ComFuncDefinition                                ' Common Digital IO Function Definitions
Global Prg_DigDef(0 To MAX_DIG_PRG) As ComFuncDefinition                                ' PurgeAir Source Digital IO Function Definitions
Global Stn_DigDef(0 To MAX_DIG_STN) As StnFuncDefinition                                ' Station Digital IO Function Definitions
Global Mfc_FunDef(0 To MAXMFC) As MfcFuncDefinition                                     ' MFC Function Definitions


' *******************************************************************************
' Reporting Configuration Data Type
' *******************************************************************************
Type ReportingConfiguration
    CsvEotReporting  As Boolean
    CsvEotSummary  As Boolean
    CsvEotDetail  As Boolean
    CsvGenReporting  As Boolean
    CsvGenSummary  As Boolean
    CsvGenDetail  As Boolean
    TextEotReporting  As Boolean
    TextEotSummary  As Boolean
    TextEotSummary_AutoPrint  As Boolean
    TextEotDetail  As Boolean
    TextGenReporting  As Boolean
    TextGenSummary  As Boolean
    TextGenDetail  As Boolean
    XlsEotReporting  As Boolean
    XlsEotSummary  As Boolean
    XlsEotDetail  As Boolean
    XlsGenReporting  As Boolean
    XlsGenSummary  As Boolean
    XlsGenDetail  As Boolean
End Type
'Global Cfg_Reporting As ReportingConfiguration
' *******************************************************************************
' LeakTest Configuration Data Type
' *******************************************************************************
Type LeakTestConfiguration
    timeOut  As Single
    PressTimeout  As Single
    PressTol  As Single
    PressTolDuration  As Single
    DeffTol  As Single
    InitialN2Flow  As Single
    ReportInterval  As Single
End Type
Global Cfg_LeakTest As LeakTestConfiguration
' *******************************************************************************
' LeakTest Recipe Data Type
' *******************************************************************************
Type LeakTestRecipe
    TargetPress  As Single
    HoldDuration  As Single
End Type
Global Rcp_LeakTest As LeakTestRecipe
' *************************************************************************************
' Declare SixteenBits Data Type
'
'
'
Type SixteenBits
    ' Actually Only 15 Bits
    B00 As Boolean
    B01 As Boolean
    B02 As Boolean
    B03 As Boolean
    B04 As Boolean
    B05 As Boolean
    B06 As Boolean
    B07 As Boolean
    B08 As Boolean
    B09 As Boolean
    B10 As Boolean
    B11 As Boolean
    B12 As Boolean
    B13 As Boolean
    B14 As Boolean
End Type

' *************************************************************************************
' Declare (LiveFuel) AutoDrainFill Definition Type
'
'
Type AdfDefinitionType
    hasLIVEFUEL As Boolean                  ' this station uses Live Fuel
    hasAUTODRAINFILL As Boolean             ' this station's LiveFuel Vapor Generator Tank supports Auto Drain/Fill
    hasADF_VaporValve As Boolean            ' this station's LiveFuel Vapor Generator Tank has a VaporValve
    hasADF_Heater As Boolean                ' this station's LiveFuel Vapor Generator Tank has an Electric Heater
    hasADF_WaterBath As Boolean             ' this station's LiveFuel Vapor Generator Tank has a WaterBath Heater/Chiller
    hasADF_PS As Boolean                    ' this station's LiveFuel Vapor Generator Tank has an N2 Pressure Switch
    hasADF_FST As Boolean                   ' this station's LiveFuel Vapor Generator Tank has a Fuel Storage Tank
    hasADF_LT As Boolean                    ' this station's LiveFuel Vapor Generator Tank has a Level Transmitter
    TankNum As Boolean                      ' this station's LiveFuel Vapor Generator Tank Number
End Type
Global AdfDef(1 To MAX_STN) As AdfDefinitionType

' *************************************************************************************


' Declare array for Station Information
'
'                   Station Types
'                   1 = Regular
'                   2 = ORVR  (Same As Regular - as of 14 March 2007)
'                   3 = Live Fuel
'                   9 = Dummy
'
'                   AutoDrainFill Live Fuel TankTypes
'                   0 = None (i.e. No I/O)                                              (first=before 2004)
'                   1 = ADF #1; Pump, Drain, Fill                                       (first=MarkIV)
'                   2 = Pump, Drain, Fill, Vapor                                        (none so far)
'                   3 = Pump, Drain, Fill, Vapor                                        (none so far)
'                   4 = Pump, Drain, Fill, Vapor                                        (none so far)
'                   5 = Pump, Drain, Fill, Vapor                                        (none so far)
'                  11 = Pump, Drain, Fill, with Heater                                  (none so far)
'                  12 = ADF #2; Pump, Drain, Fill, Vapor, Bypass, N2Purge with Heater   (first=Mahle;Jan2006)
'                  13 = Pump, Drain, Fill, with Heater                                  (none so far)
'                  14 = Pump, Drain, Fill, with Heater                                  (none so far)
'                  15 = Pump, Drain, Fill, with Heater                                  (none so far)
'                  22 = ADF #3; Pump, Drain, Fill, Vapor, No Heater, FuelStorage Tank   (first=Chrysler;June2015)
'                  90 = No ADF, just WaterBath                                          (Honda R&D; Nov 2017)
'
'                   Aspirator Number
'                   0    = use Common(i.e. default) Aspirator        (default = 1)
'                   1-19 = use Aspirator #
'
'
Type StnInfo
    desc As String                                                   ' Station Text Descriptor
    Type As Integer                                                  ' Station Type
    AspiratorNum As Integer                                          ' Aspirator Number for station
    ADF_StnNum As Integer                                            ' AutoDrainFill Unit Number for station (0=no ADF or NA)
    ADF_DEF As AdfDefinitionType                                     ' AutoDrainFill Definition
    ADF_HEATERTYPE As Integer                                        ' AutoDrainFill Heater Type for station (0=no Heater or NA, 1=Electric, 2=WaterBath)
    ADF_TANKTYPE As Integer                                          ' AutoDrainFill Tank Type for station (0=no ADF or NA; built from AdfDef bits)
    DefAuxScale As Integer                                           ' Default Aux Scale for station (0=none or NA)
    DefPriScale As Integer                                           ' Default Primary Scale for station (0=none or NA)
    ButMfcDensityMult As Single                                      ' Butane MFC Density Multiplier (0.9-1.1)
    ButMfc2DensityMult As Single                                     ' Butane ORVR2 MFC Density Multiplier (0.9-1.1)
    USINGPURGEOVEN As Boolean                                        ' Optional Oven for Purge
End Type
Global STN_INFO(1 To MAX_STN) As StnInfo                            ' Station Information Arrays
Global Def_Stn As StnInfo                                           ' Station Information Array for Station Definition
Global Def_Opto(16) As Integer                                      ' Opto Module Type Information Array for Station Definition
Global Opto_Info(0 To MAX_ADDR, 16) As Integer                      ' Opto Module Type Information Array (BaseAddr, Slot#)
Global Node_Info(0 To MAX_ADDR) As Integer                          ' Node Information Array(0=no board,8=8slot board,12=12slot board,16=16slot board)
' Const ADFtankPDF = 1
' Const ADFtankPDFVBypassHeat = 12
' Const ADFtankPDFVNoHeatStorage = 22

' *************************************************************************************
' Declare Control Block for "PurgeAir Supplies"
'
'
'
Type PrgAirInfo
  desc As String                            ' set by PurgDef; PurgeAir Supply Text Descriptor
  CheckSecs As Integer                      ' set by PurgDef(currently hardcoded); Number of seconds between successive checks of the "Request Flag"
  RequestRdy As Boolean                     ' PurgeAir Supply Ready Request Flag (from station(s)); station will want purgeair soon
  RequestRun As Boolean                     ' PurgeAir Supply Run Request Flag (from station(s)); station wants purgeair now
  LastRequestRdy As Boolean                 ' Value of "RequestRdy Flag" Last Time it was checked
  LastRequestRun As Boolean                 ' Value of "RequestRun Flag" Last Time it was checked
  StandbyRequest As Boolean                 ' PurgeAir Standby Request Flag (from station(s)); station is testing
  LastStandbyRequest As Boolean             ' Value of "Standby Request Flag" Last Time it was checked
  lastTime As Date                          ' DateTime when "Request Flag" was last checked
  StandingBy As Boolean                     ' PurgeAir Supply is StandingBy (i.e. Ready to be Requested to be Ready)
  Requested As Boolean                      ' PurgeAir Supply is Requested to be Ready
  Running As Boolean                        ' PurgeAir Supply is Running
  Ready As Boolean                          ' PurgeAir Supply Ready Flag (to station(s)); PurgeAir Supply is Ready to Run
  UsingPrgReqAK As Boolean                  ' set by PurgDef; this PurgeAir Source uses Purge Request/Ready via AK commands
  UsingPrgReqHdw As Boolean                 ' set by PurgDef; this PurgeAir Source uses Purge Request/Ready (DO/DI) Hardware
  UsingVacSwHdw As Boolean                  ' set by PurgDef; this PurgeAir Source uses Vacuum Switch(s)
  UsingAuxAirSol As Boolean                 ' set by PurgDef; this PurgeAir Source uses an Aux Air Valve
  UsingPosPrsPrg As Boolean                 ' set by PurgDef; this PurgeAir Source can perform Positive Pressure Purges
End Type
Global PRG_INFO(1 To MAX_PRG) As PrgAirInfo                         ' PurgeAir Supply Information Arrays
Global Def_Prg As PrgAirInfo                                        ' PurgeAir Supply Information Array for PurgeAir Supply Definition
' *************************************************************************************
Global LastPurgeStart As Date                ' DateTime when most Recent Purge was Started


' *************************************************************************************
' Declare Control Block for PAS Control
'
'
'
Type PAScontrol
'  Auto As Boolean                           ' PAS Local Control Mode   (1=Auto,0=Manual)
  Ok As Boolean                             ' PAS parameter has been within the limits for at least the desired duration
  Duration As Double                        ' How many seconds the PAS parameter has been within limits
  DurationTarget As Double                  ' How many seconds the PAS parameter must be within limits
  LastUpdate As Double                      ' Timer value when PAS parameter was last checked
  timeOut As Boolean                        ' PAS parameter has been outside the limits for too long
  TimeOutDuration As Double                 ' How many seconds the PAS parameter has been outside limits
  TimeOutTarget As Double                   ' How many seconds the PAS parameter must be outside limits for a Timeout
End Type
Global PAS_INFO(1 To 2) As PAScontrol       ' PAS Control Blocks       (1=Temp,2=Moist)
Global last_INFO(1 To 2) As PAScontrol      ' PAS Control Blocks       (1=Temp,2=Moist)
' *************************************************************************************


' *************************************************************************************
' Declare Control Block for PID Control (also used for On/Off Control)
'
'
'
Type PIDcontrol
  Enable As Boolean                         ' Control Mode   (0=Not Enabled ,1=Enabled (i.e. Run in Auto; if not inhibited)
  SP As Single                              ' Set Point as percent
  PV As Single                              ' Process Variable as percent
  Er As Single                              ' Error (SP - PV) as Percent
  out As Single                             ' Output as percent
  outmax As Single                          ' Max Value for Output as percent
  outmin As Single                          ' Min Value for Output as percent
  Pgain As Single                           ' Proportional Gain
  Igain As Single                           ' Integral Gain
  Dgain As Single                           ' Derivative Gain
  CumI As Single                            ' Integral Term (cumulative)
  CumImax As Single                         ' Max value for Integral Term (cumulative)
  CumImin As Single                         ' Min value for Integral Term (cumulative)
  Rev As Boolean                            ' Reverse Action (true/false)
  LastUpdate As Double                      ' Timer value when PID was last updated
  Inhibit As Boolean                        ' Output must be Off (or 0%) if Inhibit is True
  OffTimer As Double                        ' Number of Seconds Output has been Off
  OffDuty As Double                         ' Min Seconds Output must stay Off
  OffDutyMult As Single                     ' Off Duty Multiplier (for user tweaking of OffDuty)
  OffLimitDelta As Single                   ' Delta + SP = Temp above which the Heater is Off (usually near the SP)
  OnTimer As Double                         ' Number of Seconds Output has been ON
  OnDuty As Double                          ' Min Seconds Output must stay ON
  OnDutyMult As Single                      ' On Duty Multiplier (for user tweaking of OnDuty)
  OnLimitDelta As Single                    ' Delta + SP = Temp below which the Heater is On  (usually near the low OOT limit)
  Output As Boolean                         ' On/Off Output
End Type
Global PID_INFO(1 To MAX_CONTROLLER) As PIDcontrol       ' Control Data Blocks       (1=PAS Temp, 2=PAS Moisture,3-10 undefined, 11-19=STN LoadRate)
' *************************************************************************************


' *************************************************************************************
' Declare System Definition Dataset (actually data subset)
'
'
Type Sysdef
  USINGC As Boolean                     ' Temperature readings are in degrees Centigrade
  USINGF As Boolean                     ' Temperature readings are in degrees Fahrenheit
  USINGMoist_RH As Boolean              ' Moisture in grains per pound
  USINGMoist_Grains As Boolean          ' Moisture in relative humidity (percent)
  USINGLVol_Engl As Boolean             ' Line Vol Calc inputs are: ID-in, Length-feet, Vol-Liters
  USINGLVol_SI As Boolean               ' Line Vol Calc inputs are: ID-mm, Length-meters, Vol-Liters
End Type
Global SysSysDef As Sysdef                  ' System Definition Information Array
Global Gen_Sysdef As Sysdef                 ' System Definition Information Array

' *************************************************************************************
' Declare Configuration DataType
'
'
Type Configuration
    Next_File As Long
    ReportFileName1stPart As Integer
    ReportFileName2ndPart As Integer
    ReportFileName3rdPart As Integer
    AutoLogon As Integer
    AutoLogonUser As String
    RptConfig As ReportingConfiguration
'    AutoPrint As Integer
    Tol_Mix_Ratio  As Single
    LoadPressure  As Single
    ButaneMassLimit  As Single
    LoadTimeLimit  As Single
    LoLim_Load_Flow  As Single              ' Low Limit for Tolerance Checking in %
    LoLim_Purge_Flow  As Single             ' Low Limit for Tolerance Checking in %
    Tol_Load_Total  As Single
    Tol_Purge_Total  As Single
    Tol_Nit_Flow  As Single
    Tol_Btn_Flow  As Single
    Tol_ORVRNit_Flow  As Single
    Tol_ORVRBtn_Flow  As Single
    Tol_Pur_Flow  As Single
    Tol_Lfv_Flow  As Single
    Default_Interval As Integer
    Load_Interval As Integer
    Purge_Interval As Integer
    LoadTotal_Interval As Single
    PurgeTotal_Interval As Single
    Tol_Temp  As Single
    Tol_StorageLevel  As Single
    Tol_FuelLevel  As Single
    Tol_FuelTemp  As Single
    Tol_Moisture  As Single
    Tol_PurgeOvenTemp  As Single
    Tol_WaterBathTemp  As Single
    PurgeDP_HiLimit  As Single
    Temp_Target  As Single
    Moisture_Target  As Single
    Heading As String * 60
    Heading2  As String * 60
    DbFileBackup_Active  As Boolean
    DbFileBackup_Path  As String
    ReportBackup_Active  As Boolean
    ReportBackup_Path  As String
    EventRecs As Integer
    JobRecs As Integer
    LCMinDelay As Integer
    LCSetPoint  As Single
    LCTime As Integer
    PressureDecay As Single
    NitrogenPurgeTime As Integer
    DoorOpenDelay As Integer
    UPSOpenDelay As Integer
    OOTtimeDelay As Integer
    CanVent_Delay_Max As Integer
    PosPressPurge As Boolean
    DryAirPurge As Boolean
    LeakCheck_Interval As Integer
    LeakTotal_Interval As Single
    LeakCheckFailResponse As Integer
    LoadSettleTime As Single                    ' in minutes
    PurgeSettleTime As Single                   ' in minutes
    TempRhLogInterval As Single                 ' in minutes
    TempRhLogVerbose As Boolean
    BtnFlowResp As Integer
    NitFlowResp As Integer
    PurFlowResp As Integer
    AirTempResp As Integer
    AirMoistResp As Integer
    StorageLevelResp As Integer
    FuelLevelResp As Integer
    FuelTempResp As Integer
    CanVentResp As Integer
    LoadRateResp As Integer
    PurgeDpResp As Integer
    PurgeOvenResp As Integer
    WaterBathResp As Integer
    PurgeOvenBand As Single
    WaterBathControl As Integer                 ' 0=direct control, 26=Control Fuel Temp, 28=Control Vapor Temp
End Type
Global StationConfig(1 To MAX_STN, 1 To MAX_SHIFT) As Configuration     ' Station/Shift copy of System Configuration Information Array
Global SysConfig As Configuration                                       ' System Configuration Information Array


' *************************************************************************************
' Declare AutoDrainFill (live fuel tank) DataType
'
'
Type AdfCfgType
    DrainDelay As Integer              ' Configuration for # of seconds to run pump after tripping low level switch.
    DrainTimeout As Integer            ' Configuration for # of seconds for timeout of Drain operation.
    DrainShutOff As Single             ' Configuration for % Level to ShutOff Drain operation (equivalent to LoLo Level Switch).
    FillDelay As Integer               ' Configuration for # of seconds to wait after turning off the pump.
    FillTimeout As Integer             ' Configuration for # of seconds for timeout of Fill operation.
    FillShutOff As Single              ' Configuration for % Level to ShutOff Fill operation (equivalent to Hi Level Switch).
    HeaterTimeout As Integer           ' Configuration for # of MINUTES for timeout of Heater operation.
    PurgeDrainDelay As Integer         ' Configuration for # of seconds to wait after N2 PS is On (N2 Purge before Drain)
    PurgeFillDelay As Integer          ' Configuration for # of seconds to wait after N2 PS is On (N2 Purge after Fill)
    PurgeTimeout As Integer            ' Configuration for # of seconds for timeout of N2 Purge Operation
    VaporGenTankVol As Single          ' Configuration for EU Volume of Vapor Generator Tank (EU is set by LineVol Units; SI => liters, English => gallons)
    FuelStorageTankVol As Single       ' Configuration for EU Volume of Fuel Storage Tank (EU is set by LineVol Units; SI => liters, English => gallons)
    VaporGenLevelTol As Single         ' Configuration for EU Leak Rate of Vapor Generator Tank (EU is set by LineVol Units; SI => liters/hr, English => gallons/hr)
    FuelStorageLevelTol As Single      ' Configuration for EU Leak Rate of Fuel Storage Tank (EU is set by LineVol Units; SI => liters/hr, English => gallons/hr)
    FstDrainDelay As Integer           ' Fuel Storage Tank Configuration for # of seconds to run pump after tripping low level switch.
    FstDrainTimeout As Integer         ' Fuel Storage Tank Configuration for # of seconds for timeout of Drain operation.
    FstDrainShutOff As Single          ' Fuel Storage Tank Configuration for % Level to ShutOff Drain operation (equivalent to LoLo Level Switch).
    FstFillDelay As Integer            ' Fuel Storage Tank Configuration for # of seconds to wait after turning off the pump.
    FstFillTimeout As Integer          ' Fuel Storage Tank Configuration for # of seconds for timeout of Fill operation.
    FstFillShutOff As Single           ' Fuel Storage Tank Configuration for % Level to ShutOff Fill operation (equivalent to Hi Level Switch).
End Type
Global StationCfg_ADF(1 To MAX_STN, 1 To MAX_SHIFT) As AdfCfgType


' *************************************************************************************
' Declare Recipe DataType
'
'
Type Recipe
    Number As Integer
    Name As String
    desc(0 To 2) As String
    UseAnalyzer As Boolean
    UsePriScale As Boolean
    PriScaleNo As Integer
    UseAuxScale As Boolean
    AuxScaleNo As Integer
    IDLoad As Single
    LoadL As Single
    IDPurge As Single
    PurgeL As Single
    IDVent As Single
    VentL As Single
    LoadV As Single
    PurgeV As Single
    VentV As Single
    Cycles As Single
    CyclesSave As Single
    LeakCheck As Boolean
    LeakPrimary As Boolean
    LeakAux As Boolean
    StartMethod As Integer                      ' 0=Now, 1=AfterDelay, 2=AtDate
    StartDelay As Double                        ' in minutes xxx
    StartDate As Date                           ' M/D/YYYY hh:mm
    UseTstCycSetup As Boolean                   ' 0=all cycles the same, 1=use custom test cycle setup
    Purge_Liters As Single
    Purge_Flow As Single
    Purge_Can_Vol As Single
    Purge_Time As Single
    Purge_AuxTime As Single
    PurgeAuxCan As Boolean
    PurgeCansInSeries As Boolean
    Purge_Method As Integer
    Purge_ProfileNumber As Integer
    Purge_TargetMode As Integer                 ' Continuous Purge = 1; PurgePauseRepeat = 2; undefined/na = 0
    Purge_TargetWeight As Single
    Purge_TargetWC As Single
    Purge_MaxVolumes As Integer
    Purge_TargetPurge As Single
    Purge_TargetPause As Single
    PurgeOven As Boolean
    PurgeOvenSP As Single
    PauseAfterPurge As Boolean
    PauseAfterPurgeForOper As Boolean
    PausePurgeTime As Single
    Load_Method As Integer
    Load_MethodSave As Integer
    UseHiRangeMFC As Boolean
    UseButane As Boolean
    Load_Rate As Single
    Load_RateSave As Single
    UseLoadRatePID As Boolean
    WC_Mult As Single
    WC_MultSave As Single
    EPAFill As Integer
    Load_Time As Single
    Mix_Percent As Single
    Load_Wt As Single
    LoadBreakthrough As Single
    FIDmg As Single
    NitrogenFlow As Single
    NitrogenFlowSave As Single
    FuelStartTemp As Single
    FuelRampRate As Single
'    TargetRateConcentration As Single
'    TargetConcentration As Single
'    DwellTime As Single
    MaxLoadTime As Integer                      ' in minutes xxxx
    LiveFuel As Boolean
    LiveFuelChgAuto As Boolean
    LiveFuelChgFreq As Integer
    ADF_Heater As Boolean
    ADF_HeaterSP As Single
    PauseAfterLoad As Boolean
    PauseAfterLoadForOper As Boolean
    PauseLoadTime As Single
    PauseAfterLeak As Boolean
    PauseLeakTime As Single
    AuxOutputs As Boolean
    AuxOutputs_Load(1 To 4) As Boolean
    AuxOutputs_Purge(1 To 4) As Boolean
    CycleType As Integer                        ' 0=Purge/Load, 1=Load/Purge
    EndMethod As Integer                        ' 0=Cycles, 1=StableWeightChange
    EndWeightTolerance As Single                ' in grams xxx.xxx
    EndConsecutiveCycles As Integer             ' number of consecutive cycle with weight change within tolerance
    EndMaximumCycles As Integer                 ' maximum number of cycles before job must be stopped
    EndMinimumCycles As Integer                 ' minimum number of cycles before job can be complete
    UpdateCanWc As Boolean
    Validated As Boolean
End Type
Global StationRecipe(1 To MAX_STN, 1 To MAX_SHIFT) As Recipe                    ' Station/Shift Process Recipe Information Array
Global CourseRecipes(1 To MAX_STN, 1 To MAX_SHIFT, 1 To MAX_COURSES) As Recipe   ' JobSequence Process Recipe Information Array
Global EmptyRecipe As Recipe                                                    ' Blank Recipe
Global ExportedRecipe As Recipe                                                 ' Recipe Exported from the Recipe screen

' *************************************************************************************
' Declare Purge Profile DataType
'
'
Type PurgeProfileType
    Number As Integer
    Description As String
    Duration As Single                                      ' in minutes
    DurDesc As String                                       ' h,hhh:mm:ss
    ProjectedLiters As Single
    ProjectedVolumes As Single
    EndStep As Integer
    StepStartSetpoint(0 To MAX_PROFILESTEPS) As Single
    StepDuration(0 To MAX_PROFILESTEPS) As Single
    StepType(0 To MAX_PROFILESTEPS) As Integer               ' 0 = undefined; 1 = step MfcSetPoint; 2 = ramp MfcSetPoint; 3 = last Step
    Validated As Boolean
End Type
Global StationProfile(1 To MAX_STN, 1 To MAX_SHIFT) As PurgeProfileType     ' Station/Shift Purge Profile Information Array
Global ExportedProfile As PurgeProfileType

' *************************************************************************************
' Declare Canister Recipe DataType
'
'
Type CanisterRecipe
    Number As Integer
    Description As String
    WorkingCapacity As Single
    WorkingVolume As Single
    Validated As Boolean
End Type
Global StationCanister(1 To MAX_STN, 1 To MAX_SHIFT) As CanisterRecipe      ' Station/Shift Canister Information Array
Global Export_Canister As CanisterRecipe


' *************************************************************************************
' Declare Remote Task (Remote) DataType
'
'
Type RemData
    TaskID As String
    VIN As String
    Specialist As String
    TaskStatus As String
    PreviousResult As String
    ReportComplete As Boolean
    Can As CanisterRecipe
    Rcp As Recipe
    prf As PurgeProfileType
    RequestedStation As Integer
    RequestedShift As Integer
    ActualStation As Integer
    ActualShift As Integer
End Type
Global CurRemoteTask As RemData
Global StnRemoteTask(1 To MAX_STN, 1 To MAX_SHIFT) As RemData

' *************************************************************************************
' Declare Job Sequence Dataset
'
'   A Job Sequence is a sequence of up to 99 Courses
'       Each Course is one of the following:
'           1. Wait for Operator OK
'           2. Pause for x minutes
'           3. Run Recipe #n
'
Type JobSequenceCourse
    CourseNumber As Integer
    Type As Integer                 ' 0 = invalid, 1=waitforok, 2=pause, 3=recipe, >3=unused
    OkToProceed As Boolean          ' true = proceed to next step
    PauseDuration As Single         ' minutes
    RecipeNumber As Integer         ' master recipe number
    Cycles As Integer               ' number of cycles
    LoadRate As Single              ' grams/hr
    PurgeRate As Single             ' slpm
    EstCourseDuration As Single     ' Estimated Course Duration in Minutes
    MsgText As String
    DtsStart As Date
    DtsEnd As Date
End Type
Type JobSequence
    Number As Integer               ' Sequence Number   (primary key for Master Sequences)
    Description As String
    NumCourses As Integer
    PriScaleNo As Integer
    AuxScaleNo As Integer
    IDLoad As Single
    LoadL As Single
    IDPurge As Single
    PurgeL As Single
    IDVent As Single
    VentL As Single
    LoadV As Single
    PurgeV As Single
    VentV As Single
    CourseData(0 To 99) As JobSequenceCourse
    EstSeqDuration As Single        ' Estimated Sequence Duration in Minutes
    EstSeqDurDesc As String         ' Estimated Sequence Duration Description
    Validated As Boolean
End Type
Global StationSequence(0 To MAX_STN, 0 To MAX_SHIFT) As JobSequence      ' Station/Shift Job Sequence Information Array
Global Gen_Sequence As JobSequence

' *************************************************************************************
' Columns Of Data
Type ColumnsOfData
    Header1 As String
    Header2 As String
    Header3 As String
    HdrAlign As String
    DataFormat As String
    DataName As String
    DataOffset As Integer
    ColAlign As String
    ColLeft As Integer
    ColWidth As Integer
    InUse As Boolean
End Type
' *************************************************************************************

' *************************************************************************************
' Data Watcher data sets
Type LeakData
    ClkTime As Date
    Pressure As Single
    Comment As String
    TstTimr As Double
    isBlank As Boolean
End Type
Type LoadData
    ClkTime As Date
    NitFlow As Single
    BtnFlow As Single
    MixPcnt As Single
    LoadRate As Single
    loadTotalGrams As Single
    PriScle As Single
    AuxScle As Single
    WtChange As Single
    WtChgRate As Single
    LFcycls As Integer
    FuelTmp As Single
    TstTimr As Double
    isBlank As Boolean
End Type
Type PurgeData
    ClkTime As Date
    PrgFlow As Single
    PrgTemp As Single
    PrgHumd As Single
    PriScle As Single
    AuxScle As Single
    WtChange As Single
    WtChgRate As Single
    VolTotl As Single
    TstTimr As Double
    isBlank As Boolean
End Type
Global BlankLeakData As LeakData
Global GenCumLeakData As LeakData
Global StnLeakData(1 To MAX_STN, 1 To MAX_SHIFT, 0 To 9) As LeakData
Global BlankLoadData As LoadData
Global GenCumLoadData As LoadData
Global StnLoadData(1 To MAX_STN, 1 To MAX_SHIFT, 0 To 9) As LoadData
Global BlankPurgeData As PurgeData
Global GenCumPurgeData As PurgeData
Global StnPurgeData(1 To MAX_STN, 1 To MAX_SHIFT, 0 To 9) As PurgeData
' *************************************************************************************
' *************************************************************************************
Type LT2_ReportData
    ClkTime As Date
    SecTimer As Double
    NitFlow As Double   ' m2/s
    EffDia As Double    ' inches
    InPress As Double   ' kPa
    AtmPress As Double  ' kPa
    NitTemp As Double   ' degK
    isBlank As Boolean
End Type
Global BlankLT2_Data As LT2_ReportData
Global CurrLT2_Data As LT2_ReportData       ' (station, shift) ???
Global GenLT2_Data As LT2_ReportData
Global StnLT2Data(1 To MAX_STN, 1 To MAX_SHIFT, 0 To 9) As LT2_ReportData

' *************************************************************************************
' Station Detail Button Control data
Type ButtonControlData
    Top As Long
    Enabled As Boolean
    ToolTipText As String
    Visible As Boolean
End Type
Global Stn_ContinueBtn(0 To MAX_MODE) As ButtonControlData
' *************************************************************************************

' *************************************************************************************
' User Logon data
Type Passkey
  USER As String
  PWord As String
  Access As String
End Type
Global ApsUser As Passkey
Global CurrentUser As Passkey
Global MasterUser As Passkey
Global DefaultUser As Passkey
' *************************************************************************************

' *************************************************************************************
' Cycle Weight data
'
'
Type CycleWeights
  Cycle_StartWeight_Total As Single
  Cycle_EndWeight_Total As Single
  Load_StartWeight_Aux As Single
  Load_EndWeight_Aux As Single
  Load_StartWeight_Pri As Single
  Load_EndWeight_Pri As Single
  Purge_StartWeight_Aux As Single
  Purge_EndWeight_Aux As Single
  Purge_StartWeight_Pri As Single
  Purge_EndWeight_Pri As Single
  Load_TotalGrams As Single
  LoadPause_StartWeight_Aux As Single
  LoadPause_EndWeight_Aux As Single
  LoadPause_StartWeight_Pri As Single
  LoadPause_EndWeight_Pri As Single
  PurgePause_StartWeight_Aux As Single
  PurgePause_EndWeight_Aux As Single
  PurgePause_StartWeight_Pri As Single
  PurgePause_EndWeight_Pri As Single
End Type
Global StationCycleWeightData(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAX_CYCLES) As CycleWeights
'Global Gen_CycleWeights(0 To MAX_CYCLES) As CycleWeights
' *************************************************************************************

' *************************************************************************************
' Cycle Sequence Data
Type CycleSequenceData
    PurgeType As Integer
    PurgePara1 As Integer
    PurgePara2 As Integer
    PurgePara3 As Boolean
    PostPurgeDelay As Boolean
    PostPurgeTime As Integer
    LoadType As Integer
    LoadFuel As Integer
    LoadPara1 As Integer
    LoadPara2 As Integer
    LoadPara3 As Boolean
    LoadPara4 As Integer
    LoadPara5 As Integer
    PostLoadDelay As Boolean
    PostLoadTime As Integer
End Type
'Global Stn_CycData(1 To MAX_STN, 1 To MAX_CYCLES) As CycleSequenceData
'Global Rcp_CycData(1 To MAX_RCP, 1 To MAX_CYCLES) As CycleSequenceData
' *************************************************************************************


' *************************************************************************************
' Declare LeakCheck Control Block DataType
'
'
Type LeakCheckControlBlockType
    CycleStartRequest As Boolean
    Method As Integer                   ' 0=idle, 1=primary, 2=aux
    StartTime As Date
    StartTimer As Double
    Phase As Integer
    PhaseDts As Date
    PhaseStartDts As Date
    PhaseStartTimer As Double
    ElapsedHours As Single              ' Elapsed time leakchecking in hours
    ElapsedHours_Prev As Single         ' Previous(i.e. Before OOT or Alarm) Elapsed time leakchecking in hours
    station As Integer                  ' which station is controlling the leakcheck resources
    Shift As Integer                    ' which shift is controlling the leakcheck resources
End Type
Global LeakCheckControl As LeakCheckControlBlockType     ' Station/Shift LeakCheck Control Block Array


' *************************************************************************************
' Declare LeakTest Control Block DataType
'
'
Type LeakTestControlBlockType
    CycleStartRequest As Boolean
    TaskDesc As String
    StepDesc As String
    Step As Integer
    StartTime As Date
    StartTimer As Double
    Phase As Integer
    PhaseDts As Date
    PhaseStartDts As Date
    PhaseStartTimer As Double
    ElapsedSeconds As Single              ' Elapsed time leaktesting in seconds
    ElapsedSeconds_Prev As Single         ' Previous(i.e. Before OOT or Alarm) Elapsed time leaktesting in seconds
End Type
Global LeakTestControl(1 To MAX_STN, 1 To MAX_SHIFT) As LeakTestControlBlockType     ' Station/Shift LeakCheck Control Block Array


' *************************************************************************************
' Declare Load Control Block DataType
'
'
Type LoadControlBlockType
    CycleStartRequest As Boolean
    Method As Integer
    StartTime As Date
    StartTimer As Double
    Phase As Integer
    PhaseDts As Date
    PhaseStartDts As Date
    PhaseStartTimer As Double
    ElapsedStartDts As Date
    ConcordanceIsOpen As Boolean        ' Concordance screen is Open
    NetWtChgIsOpen As Boolean           ' NetWtChg screen is Open
    LoadTarget As Single
    LoadRateTarget As Single            ' Target LoadRate in grams/hour
    LoadRate As Single                  ' Grams per Hour Loaded (thru Butane MFC for Butane; Rate of TotalWtChg for LiveFuel)
    loadTotalGrams As Single            ' Total Grams Loaded thru MFC (Butane for Butane; TotalWtChg for LiveFuel
    LoadTotalLiters As Single           ' Total Liters Loaded thru MFC (Nitrogen for Butane; VaporCarrier for LiveFuel
    ElapsedHours As Single              ' Elapsed time loading in hours
    ElapsedHours_Prev As Single         ' Previous(i.e. Before OOT or Alarm) Elapsed time leakchecking in hours
    CurrLoadDensity As Single           ' Current Grams/Liter Loaded thru MFC (Butane for Butane; VaporCarrier for LiveFuel
    CurrWtChgRate As Single             ' Primary+Aux Scale Rate of Weight Change in grams/hour; Recently
    TotalWtChgRate As Single            ' Primary+Aux Scale Rate of Weight Change in grams/hour; for entire Load
    TotalWtChg As Single                ' Primary+Aux Scale Weight Change in grams; for entire Load
    AuxWtChgAtEOL As Single             ' Aux Scale Weight Change in grams; for entire Load as measured at End-Of-Load
    PriWtChgAtEOL As Single             ' Primary Scale Weight Change in grams; for entire Load as measured at End-Of-Load
    TotalWtChgAtEOL As Single           ' Primary+Aux Scale Weight Change in grams; for entire Load as measured at End-Of-Load
    PriWtChg As Single                  ' Primary Scale Weight Change in grams; for entire Load
    AuxWtChg As Single                  ' Aux Scale Weight Change in grams; for entire Load
    AuxWt_Start As Single               ' Aux Scale Weight at the Start of a Load Cycle
    AuxWt_End As Single                 ' Aux Scale Weight at the End of a Load Cycle
    PriWt_Start As Single               ' Primary Scale Weight at the Start of a Load Cycle
    PriWt_End As Single                 ' Primary Scale Weight at the End of a Load Cycle
    WC_Load_Time As Date                ' Working capacity load time
    WC_Load_Rate As Single              ' Working capacity load rate
    WaterBathPV As Single
    WaterBathSP As Single
    WaterBathTempOK As Boolean
End Type
Global LoadControl(1 To MAX_STN, 1 To MAX_SHIFT) As LoadControlBlockType     ' Station/Shift Load Control Block Array

' *************************************************************************************
' Declare Purge Control Block DataType
'
'
Type PurgeControlBlockType
    CycleStartRequest As Boolean
    Method As Integer
    StartTime As Date
    StartTimer As Double
    Phase As Integer
    PhaseDts As Date
    PhaseStartDts As Date
    PhaseStartTimer As Double
    Purge_Target As Single
    Purge_Total As Single               ' cumulative in liters
    Purge_Volumes As Single             ' cumulative in canister volumes
    ElapsedHours As Single              ' Elapsed time loading in hours
    ElapsedHours_Prev As Single         ' Previous(i.e. Before OOT or Alarm) Elapsed time leakchecking in hours
    CurrWtChgRate As Single             ' Primary+Aux Scale Rate of Weight Change in grams/hour; Recently
    TotalWtChgRate As Single            ' Primary+Aux Scale Rate of Weight Change in grams/hour; for entire Purge
    TotalWtChg As Single                ' Primary+Aux Scale Weight Change in grams; for entire Purge
    TotalWtChgAtEOP As Single           ' Primary+Aux Scale Weight Change in grams; for entire Purge as measured at End-Of-Purge
    AuxWtChg As Single                  ' Aux Scale Weight Change in grams; for entire Purge
    PriWtChg As Single                  ' Pri Scale Weight Change in grams; for entire Purge
    AuxWt_Start As Single               ' Aux Scale Weight at the Start of Purge Cycle
    AuxWt_End As Single                 ' Aux Scale Weight at the End of Purge Cycle
    PriWt_Start As Single               ' Primary Scale Weight at the Start of Purge Cycle
    PriWt_End As Single                 ' Primary Scale Weight at the End of Purge Cycle
    maxStep As Integer
    curCycle As Integer
    curStep As Integer                  ' when Purge-To-Target: 0=Pause; 1=Purge
    CurMfcSp As Single
    StepStartDTS As Date
    StepElapsedMinutes As Single
    StepElapsedSeconds As Long
    ProfileStartDTS As Date
    ProfileElapsedMinutes As Single     ' since start of purge
    ProfileElapsedSeconds As Long       ' since start of purge
    CompletedStepMinutes As Single      ' cumulative (not counting current step)
    PurgeOvenPV As Single
    PurgeOvenSP As Single
    PurgeOvenTempOK As Boolean
    PurgeTargetWt As Single             ' in grams
    InhibitOotCheck As Boolean
End Type
Global PurgeControl(1 To MAX_STN, 1 To MAX_SHIFT) As PurgeControlBlockType     ' Station/Shift Purge Control Block Array

' *************************************************************************************
' Declare Job Info DataType
'
'
Type JobInfoBlockType
    Comment As String                ' Comments Field
    End_Op As String                 ' Ending Operator
    Engineer As String               ' Engineer Data field
    Start_Op As String               ' Start Operator
    Vehicle As String                ' Vehicle ID Data Field
    Start_Baro As Single             ' Barometer value Start
    End_Baro As Single               ' End
    End_OK As Boolean                   ' Used to indicate if Station terminated normally
End Type
Global JobInfo(1 To MAX_STN, 1 To MAX_SHIFT) As JobInfoBlockType     ' Station/Shift Job Information Array

' *************************************************************************************
' Declare Station Control Block DataType
'
'
Type StationControlBlockType
    AbortRequest As Boolean
    StartRequest As Boolean
    ContinueRequest As Boolean
    StopRequest As Boolean
    Actual As Single
    ActualAtEnd As Single
    Target As Single
    Course As Integer
    CurrCycle As Integer
    Mode As Integer
    Mode_Last As Integer
    Mode_PauseSave As Integer
    Mode_StartDts As Date
    ModeIsIdle_DebounceCount As Integer
    ModeIsIdle_Debounced As Boolean
    TestTimer As Double
    TestTimerIsRunning As Boolean
    BtnDensity As Single
    Scale_OK As Boolean
    ScalesInUse As Boolean
    AuxScaleStn As Single               ' Station that "Owns" the Aux Scale;        Owns = has the valves
    PriScaleStn As Single               ' Station that "Owns" the Primary Scale;    Owns = has the valves
    AuxScaleWt As Single                ' Aux Scale Weight in grams
    PriScaleWt As Single                ' Primary Scale Weight in grams
    AuxTare As Single                   ' Aux Scale Tare weight
    PriTare As Single                   ' Primary Scale Tare weight
    AuxWt_Start As Single               ' Aux Scale Weight at the Start of Purge Cycle
    AuxWt_End As Single                 ' Aux Scale Weight at the End of Purge Cycle
    PriWt_Start As Single               ' Primary Scale Weight at the Start of Purge Cycle
    PriWt_End As Single                 ' Primary Scale Weight at the End of Purge Cycle
    Job_Number As String * 6            ' 6 char Job Number
    Job_Description As String * 50      ' 50 char Job Description
    StartMethod As Integer
    EndMethod As Integer
    EstJobDur As Single                 ' Estimated Job Duration in Minutes
    EstJobDurDesc As String             ' Estimated Job Duration Description
    DelaySeconds As Double
    DelayToGo As Double
    NewDataInDB As Boolean
    DBFile As String                    ' Database File Name
    RptFile As String                   ' Report File Name Kernel
    Start_Time As Date
    End_Time As Date
    End_Timer As Double
    OotCurrent As Integer               ' current OOT
    OotResponse As Integer              ' Response to current OOT
    PausedDts As Date
    IsPausedInAlarm As Boolean
    PauseAlarmStartTime As Date         ' When pause occured
    AlarmDelayTime As Long              ' How many seconds delayed now
    PauseMessage As String
    LeakCheckStatus As Integer          ' Of Primary Canister; 0=unknown
    LcStatusDescription As String       ' Description of LeakCheck Status of Pri Canister
    LiveFuelCycleCount As Integer       ' Count of Cycles since last Live Fuel change
    CompletedCycles As Integer
    CompletedLoads As Integer
    CompletedPurges As Integer
End Type
Global StationControl(1 To MAX_STN, 1 To MAX_SHIFT) As StationControlBlockType     ' Station/Shift Control Block Array


' *************************************************************************************
' Declare (LiveFuel) AutoDrainFill Control Block Type
'
'
Type AdfControlBlockType
    AdfDefinition As AdfDefinitionType
    TurnHeaterOn As Boolean                 ' "Request" from the ADF_Heater sub to the ADF_Sequence sub to turn the LiveFuel Tank Heater On
    HeaterOff As Integer                    ' Number of Seconds the LiveFuel Tank Heater has been Off
    HeaterOn As Integer                     ' Number of Seconds the LiveFuel Tank Heater has been On
    TempinTol As Integer                    ' Number of Seconds the LiveFuel Tank Temp is in the Temp Tolerance Band
    TempOK As Boolean                       ' LiveFuel Tank Temp is Steady and in the Temp Tolerance Band
    Mode As Integer                         ' 1=DrainOnly; 2=DrainThenFill.
    Step As Integer                         ' Current Step Number of AutoDrainFill Sequence.
    StepBeforePause As Integer              ' Step Number of AutoDrainFill Sequence before Pause.
    Task As String                          ' Description of ADF_Mode
    Message As String                       ' Description of Current Step
    Step_Time As Date                       ' Max Time Allowed to Complete Current Step
    FillRequestOn As Boolean                ' "Request" from the FuelStorage Tank to the Facility Supply to fill the Fuel Storage Tank
    ButtonVisible_Done As Boolean           ' OK for ADF Display DONE button to be Visible
    ButtonVisible_Ignore As Boolean         ' OK for ADF Display IGNORE button to be Visible
    ButtonVisible_Retry As Boolean          ' OK for ADF Display RETRY button to be Visible
    ButtonVisible_Stop As Boolean           ' OK for ADF Display STOP button to be Visible
    Enable As Boolean                       ' OK to Run an ADF Sequence
    Heater_Enable As Boolean                ' OK for Auto Live Fuel Temp Control
    ManScreen_Enable As Boolean             ' OK to Display Manual Live Fuel Refill Screen
    InitialFill_Complete As Boolean         ' Initial Drain/Fill Sequence is Complete
    ReadyForLoad As Boolean                 ' OK to Run a Load Cycle; Live Fuel is Ready
    ReadyForRefill As Boolean               ' Need to Refill the Live Fuel Tank
    RefillRequest As Boolean                ' Request a Refill of the Live Fuel Tank
    SetOkRequest As Boolean                 ' Request a set of the LiveFuel State to "OK"
    LiveFuel As Boolean
    LiveFuelChgAuto As Boolean
    LiveFuelChgFreq As Integer
    LiveFuelState As Integer                ' LiveFuel State (0=Dead, 1=Weak, 2=<OK)
    LiveFuelDensityOkCnt As Integer
    LiveFuelDensityDeadCnt As Integer
    LiveFuelDensityWeakCnt As Integer
    Heater As Boolean
    HeaterSP As Single
    LevelSP As Single
End Type
Global AdfControl(1 To MAX_STN) As AdfControlBlockType
Global ADF_HeaterCheckTime As Date

' *************************************************************************************
' Declare (LiveFuel AutoDrainFill) FuelStorageTank Control Block Type
'
'
Type FstControlBlockType
    Mode As Integer                         ' 1=Drain; 2=Fill.
    Step As Integer                         ' Current Step Number of FuelStorageTank Drain & Fill Sequence.
    StepBeforePause As Integer              ' Step Number of FuelStorageTank Drain & Fill Sequence before Pause.
    Task As String                          ' Description of Current Mode
    Message As String                       ' Description of Current Step
    Step_Time As Date                       ' Max Time Allowed to Complete Current Step
    FillRequestOn As Boolean                ' "Request" from the FuelStorage Tank to the Facility Supply to fill the Fuel Storage Tank
    ButtonVisible_Fill As Boolean           ' OK for FST Display "+" button to be Visible
    ButtonVisible_Drain As Boolean          ' OK for FST Display "-" button to be Visible
    ButtonVisible_Stop As Boolean           ' OK for FST Display STOP button to be Visible
    Enable As Boolean                       ' OK to Run an FST Drain & Fill Sequence
    LevelSP As Single
End Type
Global FstControl(1 To MAX_STN) As FstControlBlockType

' *************************************************************************************
' Declare OOT Control Block DataType
'
'
Type OOTControlBlockType
    NitFlowOOT As Boolean
    BtnFlowOOT As Boolean
    PurFlowOOT As Boolean
    FuelLevelOOT As Boolean
    StorageLevelOOT As Boolean
    FuelTempOOT As Boolean
    AirTempOOT As Boolean
    AirMoistOOT As Boolean
    CanVentOOT As Boolean
    LoadRateOOT As Boolean
    PurgeDpOOT As Boolean
    PurgeOvenOOT As Boolean
    WaterBathOOT As Boolean
    NitFlowOOTCnt As Integer
    BtnFlowOOTCnt As Integer
    PurFlowOOTCnt As Integer
    FuelLevelOOTCnt As Integer
    StorageLevelOOTCnt As Integer
    FuelTempOOTCnt As Integer
    AirTempOOTCnt As Integer
    AirMoistOOTCnt As Integer
    CanVentOOTCnt As Integer
    CanVent_DelayCount As Long          ' Number of seconds so far in Load Cycle
    CanVent_DelayOn As Boolean          ' True = Override Canventalarm flow sw. contact
    CanVent_TimeNow As Date             ' Time of current logic scan
    CanVent_TimeLast As Date            ' Time of last logic scan
    CanVent_TimeDelta As Long           ' Difference between ...Time_Now and ...Time_Last
    LoadRateOOTCnt As Integer
    PurgeDpOOTCnt As Integer
    PurgeOvenOOTCnt As Integer
    WaterBathOOTCnt As Integer
End Type
Global OOTs(1 To MAX_STN, 1 To MAX_SHIFT) As OOTControlBlockType     ' Station/Shift OOT Control Block Array

' *************************************************************************************
' Declare FuelSupply DataType
'
'
Type FuelSupplyBlockType
    Date As String
    CurrentOnHand As Single
    CylinderWeight As Single
    WarningActive As Boolean
    WarningSetPoint As Single
End Type
'Butane Available
Global ButaneSupply As FuelSupplyBlockType
'Fuel Available
Global FuelSupply As FuelSupplyBlockType

' *************************************************************************************
' Declare System Timer Data DataType
'
'
Type SystTimerDataType
' SYSTEM TIMER INFORMATION
    desc As String           ' System Timer Description
    Phase As Integer         ' System Timer Current Phase
    Step As Integer          ' System Timer Current Step
    Interval As Integer      ' System Timer Interval Setting
    Actual As Double         ' System Timer Actual (Now-LastTime)
    delta As Double          ' System Timer Difference between Value & Actual
    LastTimer As Double       ' System Timer Last Timer
    max As Double            ' System Timer Max Delta
    Min As Double            ' System Timer Min Delta
End Type
Global SystemTimers(1 To 9) As SystTimerDataType

' *************************************************************************************
' Declare values for Heater Operation
'
Type ChillerInfo
    RunChiller As Boolean
    BufferIn As String
    BufferOut As String
    CurCmdIdx As Integer
    CurCmdDesc As String
    CurCmdChars As String
    CurCmdComplete As Boolean
    CurCmdTimeout As Boolean
    CmdTimeoutTimer As Single
    CmdToBeSentFlag As Boolean
    CmdSentFlag As Boolean
    CmdToBeAckFlag As Boolean
    CmdRecAckFlag As Boolean
    CmdRecChars As String
    CmdRecErrorFlag As Boolean
    CmdRecErrorNumber As Integer
    CmdRecValueChars As String
    ChillerPhase As Integer
    ChillerRunning As Boolean
    InitComplete As Boolean
    CommOK As Boolean
    CommOnline As Boolean
    ErrorCount As Integer
    MaxErrorCount As Integer
    TimeoutValue As Single
    StatusIn As Integer
    StatIn As String
    Overtemp As Boolean
    LowLevel As Boolean
    PumpBlocked As Boolean
    IntFaultMc1 As Boolean
    IntFaultMc2 As Boolean
    Type As Variant
    Version As Variant
    PvIn As Single
    SpIn As Single
    OutIn As Integer
    OperModeIn As Integer
    OvrTmpSpIn As Single
    P_In As Single
    I_In As Single
    ModeIn As Integer
    SpOut As Single
    OutOut As Integer
    OperModeOut As Integer
    P_Out As Single
    I_Out As Single
    ModeOut As Integer
End Type
Global LF_Chiller As ChillerInfo                    ' Heater/Chiller Information Array

' *************************************************************************************
' Declare Control Block for AK Remote Control
'
'
'
Type AK_ControlBlock
    Active As Boolean                           ' AK remote control is active
    Run As Boolean                              ' Test is running under AK remote control
    TestSelected(1 To 4) As Boolean             ' true = Test is selected
    PhaseSelected(1 To 4, 1 To 5) As Boolean    ' true = Test,Phase is selected
    MfcSP(1 To 4, 1 To 2) As Single             ' MFC SetPoint by (Test,MFC#) (Direct or Proportional)
    MfcSPfromRecipe(1 To 4, 1 To 2) As Single   ' MFC SetPoint by (Test,MFC#) (Direct or Proportional)
    EvacPurgeSelected As Boolean
    LeakCheckSelected As Boolean
    PurgeSelected As Boolean
    CVSflow(1 To 4) As Double
    CVSflowInitialVal(1 To 4) As Double         ' Initial CVS Flow Rate Value
    CVSflowNormalized(1 To 4) As Double         ' CVS flow as a fraction of the CVS Initial Flow Rate
End Type
Global AK_RemCntrl As AK_ControlBlock
' *************************************************************************************

' *************************************************************************************
' Declare Control Block for AK Command Code Definition
'
'
'
Type AK_CommandCode
    cmdCode As String                         ' 4 character command code
    MaxNumParams As Integer                   ' max parameters to be appended to the command
    MinNumParams As Integer                   ' min parameters to be appended to the command
    ParamsType As Integer                     ' 0 = Floating Point, 1 = Integer, 2 = string
    Available As Boolean                          ' available on this machine
End Type
Global AK_Commands(1 To 100) As AK_CommandCode
' *************************************************************************************

' *************************************************************************************
' Declare Control Block for AK Command
'
'
'
Type AKcommand
    Printable As String                       ' command string minus <STX><separater> & <ETX>
    cmdCode As String                         ' 4 character command code
    FUnumber As Integer                       ' Functional Unit Number (0 = all or na)
    NumParams As Integer                      ' How many parameters are appended to the command
    Param(1 To 10) As Double                  ' Array of parameter values
    CmdType As String * 1                     ' the first character of the printable command string (S, A or E)
    CmdIndex As Integer                       ' index into the AK command definition array
    CmdRead As Boolean                        ' the command has been interpreted
    CmdValid As Boolean                       ' the command is valid  (i.e. no data failures)
    CmdAvailable As Boolean                   ' the command is available on this system  (i.e. it is not NotAvailable)
    CmdAccepted As Boolean                    ' the command is accepted (i.e. the server system is not BUsy)
    CmdRcvdDTS As Date                        ' DTS when command was received
    CmdRcvdTimer As Double                    ' Timer value when command was received
End Type
Global AKcmd_Current As AKcommand
Global AKcmd_Last As AKcommand
' *************************************************************************************

' *************************************************************************************
' Declare Control Block for AK Response
'
'
'
Type AKresponse
    Printable As String                       ' command string minus <STX><separater> & <ETX>
    RspCode As String                         ' 4 character command response code
    ErrorStatus As Integer                    ' Error Status Number (0-9; 0 = none)
    NumParams As Integer                      ' How many parameters are appended to the response
    ParamsType As Integer                     ' 0 = Floating Point, 1 = Integer, 2 = string
    ParamNum(1 To 100) As Double               ' Array of numeric parameter values
    ParamStr(1 To 100) As String               ' Array of string parameter values
    RspSent As Boolean                        ' This response has been sent
End Type
Global AKrsp_Current As AKresponse
Global AKrsp_Last As AKresponse
' *************************************************************************************

' *************************************************************************************
' Declare Control Block for Local PurgeAirGenerator Type
'
'
'
Type LocalPAGdef
    Type As Integer                           ' 0 = None, 1 = Stand-Alone, 2 = AK Master, 3 = AK Client
    ReqIn As Boolean                          ' Request received from 'below'
    ReqOut As Boolean                         ' Request to be sent to 'above'
End Type
Global LocalPagControl As LocalPAGdef
Global PAGtypeDesc(0 To 9) As String
' *************************************************************************************

' Declare Control Block for Master PurgeAirGenerator Data
'
'
'
Type MstrPAGdat
    Status As String                          ' 4 char Master PAG Control Mode
    Temperature As Single                     ' PAG Air Temperature in deg C
    Humidity As Single                        ' PAG Air Relative Humidity in rH%
    Moisture As Single                        ' PAG Air Moisture in grains/lb
    TempSP As Single                          ' PAG Air Temperature Target in deg C
    MoistSP As Single                         ' PAG Air Moisture Target in grains/lb
    TempTol As Single                         ' PAG Air Temperature Tolerance in deg C
    MoistTol As Single                        ' PAG Air Moisture Tolerance in grains/lb
    ReqIn As Boolean                          ' Request received from 'below'
    RdyOut As Boolean                         ' Ready Output
End Type
Global MasterPagData As MstrPAGdat
' *************************************************************************************
' *************************************************************************************
' Declare Error Status Tracking
'
'
Type ErrorStatus
    AnyError As Boolean
    TempOOT As Boolean
    TempTO As Boolean
    MoistOOT As Boolean
    MoistTO As Boolean
    TestBit As Boolean
End Type
Global ErrorStatus_Current As ErrorStatus
Global ErrorStatus_Last As ErrorStatus
Global ErrorStatus_NoErrors As ErrorStatus
Global ErrorStatus_ScratchPad As ErrorStatus
Global ErrorValue_Current As Integer
Global ErrorTestBit As Boolean
' *************************************************************************************
' *************************************************************************************
' Out of tolerance check
Global LastOOTCheckTime As Date
' Station Process Variables
Global Stn_ActiveShift(1 To MAX_STN) As Integer
Global Stn_MfcSpIsSet(1 To MAX_STN) As Boolean

Global Stn_Nit_Flow_PV(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_Nit_FlowSP(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_Btn_Flow_PV(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_Btn_FlowSP(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_CommonTC(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Stn_LoadEql_StartTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_LoadEql_StartAuxWt(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_LoadEql_StartPriWt(1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global Stn_LoadEql_StartLoadTotal(1 To MAX_STN, 1 To MAX_SHIFT) As Single

' *************** Process Variables Developed within PC **************
Global Stn_Leak_Log_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_LT_Log_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_Load_Log_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_LoadLimitStartTime(1 To MAX_STN, 1 To MAX_SHIFT) As Date
Global Stn_Default_Log_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_Purge_Log_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_UseTC(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global FirstTime(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean                 ' made global  2 Mar 2005 MMW
Global Stn_AuxScale_RefValues(1 To MAX_STN, 1 To MAX_SHIFT) As Single     ' Aux Scale Reference Value (for XY graphs)
Global Stn_PriScale_RefValues(1 To MAX_STN, 1 To MAX_SHIFT) As Single     ' Primary Scale Reference Value (for XY graphs)

' *************************************************************************************
' *************************************************************************************
' *************************************************************************************

Global DISPDELAY As Variant
Global OptoDelay As Double
Global INTRODELAY As Integer
Global PAUSEDELAY As Integer

Global DefCanVol2CanWcMult As Single

Global cOFF As Integer                       ' = 0
Global cON As Integer                        ' = 1
Global cNO As Integer                        ' = 0
Global cYES As Integer                       ' = 1

Global FORMFEED As String

Global DEFAULTFILEFONT As New StdFont
Global FILEFONT As New StdFont               ' Standard windows fixed font for Detail & Summary Reports
Global FILEFONTSIZE As Integer               ' Size of text / default

Global DEFAULTREPORTFONT As New StdFont
Global REPORTFONT As New StdFont             ' Standard windows proportional font for Other Reports

Global AccessDbFileExt As String
Global AccessVerDesc As String
Global AccessVerNum As Single
Global DefaultDbFileExt As String

Global XLvernum As Double
Global XLverstr As String
Global XLverdesc As String
Global XLnumrec As Long
Global XLrecsPerSheet As Long
Global ReportsXlsFileName As String
Global ReportsXlsFileExt As String

Global ReportGenRevLvl As String             ' Revision Level Required of Report Generator Program
Global CfgRevLvl As Integer                  ' Current Revision Level for config, sysdef, etc. files

'Error Message Variables
Global ErrModule(0 To 100) As Integer
Global ErrLevel(0 To 100) As Integer

' Startup Variables
Global IntroDone As Boolean                  ' Program Intro Done
Global VarInitDone As Boolean                ' Variable Init Done
Global ReadyToRun As Boolean                 ' OK to Cycle the Program
Global UserLoginOk As Boolean                ' User Login is OK to Start the Program
Global ShuttingDown As Boolean               ' Program is Shutting Down
Global About_Counter As Integer              ' Number of Seconds the About Screen has been open (startup only)
Global AutoLogon  As Integer                 ' 0 = No Auto Logon; 1 = user; 2 = cps; 3 = aps; 4 = ApsUser

Global IoComOn  As Boolean                   ' True = Run with I/O comm, False = testing without IO
Global SclComOn  As Boolean                  ' True = Run with Scale Port comm, False = testing without reading Scales
Global ChillComOn  As Boolean                ' True = Run with Chiller comm, False = testing without Chiller
Global UseLocalErrorHandler  As Boolean      ' True = Normal operation, False = Doing debugging
Global NotDebugADF  As Boolean               ' True = Normal operation, False = Doing debugging
Global NotDebugMMW  As Boolean               ' True = Normal operation, False = Doing debugging
Global NotDebugPAS  As Boolean               ' True = Normal operation, False = Doing debugging
Global NotDebugPURGE  As Boolean             ' True = Normal operation, False = Doing debugging
Global NotDebugSCALES  As Boolean            ' True = Normal operation, False = Doing debugging
Global NotDebugREM  As Boolean               ' True = Normal operation, False = Doing debugging
Global NotDebugProf  As Boolean              ' True = Normal operation, False = Doing debugging

Global DebugAIRLOG  As Boolean               ' True = Doing debugging, False = Normal operation

Global OPTOCOM_PORT As Integer               ' chg to COMMON_INFO; opto AC24 = com2 ; using sealevel = com5 ; using Moxa = com6
Global NR_PRGAIR As Integer                  ' Number of PurgeAir Supplies
Global NR_SCALES As Integer                  ' Number of Scales (includes # of remote scales)
Global NR_SHIFT As Integer                   ' Number of shifts
Global NR_STN As Integer                     ' Number of stations
Global NR_DUMMYSTN As Integer                ' Number of dummy stations (i.e. io board only)
Global NR_REMOTESCALES As Integer            ' Number of Remote Scales (i.e. not mounted in the Station Cabinet); 0 = None
Global NR_CAN As Integer                     ' Number of Master Canisters
Global NR_RCP As Integer                     ' Number of Master Recipes
Global NR_JOBSEQ As Integer                  ' Number of Master Job Sequences
Global NR_PRGPROF As Integer                 ' Number of Master PurgeProfiles

Global LoadMfcDelayTime As Integer           ' in seconds
Global LoadEqlDelayTime As Integer           ' In seconds
Global MFC_Settle_Time As Integer            ' In seconds
Global MSGDELAY As Integer                   ' How long the DelayBox stays open; In milliseconds
Global GramsPerLiter As Single               ' default = 2.40633
Global MfcSpMin As Single                    ' default = 5.0 %
Global DefScaleMax                           ' default scale max reading (min = 0; max = 100 to 100,000)
Global MaxSheathTempForAdfDrain As Single    ' Max Sheath Temp For ADF Drain of LiveFuel Vapor Tank in degC
Global DeadLiveFuelDensity As Single         ' density of "Dead" LiveFuel Vapor in grams/liter
Global WeakLiveFuelDensity As Single         ' density of "Weak" LiveFuel Vapor in grams/liter
Global WaterBathSP As Single                 ' SetPoint to Be sent to the WaterBath

Global USINGC As Boolean                     ' Cent. readings on temp at start up (one only)
Global USINGCOMMONTC As Boolean              ' Six common thermocouples (Expert so far)
Global USINGDOOROPEN As Boolean              ' Door open switch installed
Global USINGHARDPIPEDSCALES  As Boolean      ' Scales are hard piped at 2 scales per station; stn#1 pri = 1, stn#1 aux = 2, etc.
Global USINGF As Boolean                     ' Fht readings on temp at start up   (one only)
Global USINGFIDANALYZER As Boolean           ' Analyzer (Normal=False)(Expert first user)
Global USINGFUELLEVELOOT As Boolean          ' Enable Monitor LiveFuel Level (for leaks in tank)
Global USINGLVol_Engl As Boolean             ' Line Vol Calc inputs are: ID-in, Length-feet
Global USINGLVol_SI As Boolean               ' Line Vol Calc inputs are: ID-mm, Length-meters
Global USINGMoist_Grains As Boolean          ' Moisture in relative humidity (percent)
Global USINGMoist_RH As Boolean              ' Moisture in grains per pound
Global USING_EXT_CONTACTS As Boolean         ' External contacts - Remote alarm Pause
Global DESC_EXT_CONTACTS As String           ' Description of External alarm

Global USING_ESTOP_INPUT As Boolean          ' ESTOP(12vdc or Dry Contact) Input
Global LogTempRh As Boolean                  ' Log Air Temperature and Humidity

Global USINGPASLOCALCONTROL As Boolean       ' this cps system controls the PurgeAir Generator
Global PAGSERVERIP As String                 ' IP address of this Client's AK Server

Global USING_AUX_OUTPUTS As Boolean          ' Aux (12vdc or Dry Contact) Outputs
Global NR_AUX_OUTPUTS As Integer             ' Number of Aux Outputs/Station
Global DESC_AUX_OUTPUT1 As String            ' Description of Aux Output #1
Global DESC_AUX_OUTPUT2 As String            ' Description of Aux Output #2
Global DESC_AUX_OUTPUT3 As String            ' Description of Aux Output #3
Global DESC_AUX_OUTPUT4 As String            ' Description of Aux Output #4

Global USINGSTNTC As Boolean                 ' Use 2 Thermocouple  per station
Global USINGBUTANEMASSLIMIT As Boolean       ' True is for (was Toyota only) Limit Exceeded...Aborting
Global USINGLOADTIMELIMIT As Boolean         ' For Mitsu abort by operator entered time on recipe "MAX LOAD TIME"
Global USINGCUSTOMERLOWGAS As Boolean        ' Remote inputs (was Toyota only)
Global USINGCANVENTALARM As Boolean          ' Carb #1 to use (Hardware required in opto stn)
Global CANVENTALARMDELAY As Integer          ' Delay in Minutes before checking CanVentAlarm switch
Global LAST_STN  As Integer                  ' Station Number of the Last Regular (i.e. Not Dummy) Station
Global FIRST_REMOTESCALE  As Integer         ' Scale Number of the First Remote Scale

Global USINGLINEVOLUME  As Boolean           ' Volume compensation for line lengths (CARB first user)
Global USINGOOTPAUSE  As Boolean             ' Allow Pause of station/shift on OOT condition (Config option can still turn it off)

Global USINGREMCANLOAD As Boolean
Global REM_2Gm_Recipe As Integer             ' REMOTE - Default 2 gram breakthrough recipe
Global REM_Wcm_Recipe As Integer             ' REMOTE - Default Working Capacity Multiplier recipe

Global USINGSYSTEMVACSW As Boolean
Global USINGAUXLEAKCHECK As Boolean
Global USINGLEAKCHECKEXHAUSTSOL As Boolean
Global USINGLOADPRESSURE As Boolean
Global USINGUPS  As Integer
Global USINGPRESSUREPURGE As Boolean
Global USINGDRYPURGEAIR As Boolean
Global USINGPURGEDP As Boolean
Global USINGPURGEOVEN As Boolean
Global USINGPURGESERIES As Boolean
Global USINGCONTAFTERLCFAIL As Boolean
Global USINGWATERBATH As Boolean

Global MinDataLogSeconds As Single
Global MAXNOTSTABLECOUNT As Integer
Global WEIGHTQUEUESIZE As Integer               ' # of elements in the scale weight readings queue

Global PRINTERAVAILABLE As Boolean
Global PRINTERNAME As String

Global USINGSIMULATION As Boolean               ' simple simulation(PV's set to SP's; set fixed common values; only runs if ScanIO is Off))
Global USINGSIMNOISE As Boolean                 ' add random noise (max 1 %) to simulated eu values

' chiller
Global Chiller_PORT As Integer                   ' waterbath heater rs232 com port; 9600,8,None,1
Global Chiller_Timeout As Integer                ' waterbath heater rs232 com port timeout in msec
Global ChillerCommands(0 To 7) As String
Global ChillerResponse(0 To 7) As String

Global AK_ModeDesc(0 To 99) As String           ' description of mode for use with AK Command Interface
Global AK_anychar As String                     ' "any character" character for AK commands         (default = chr(32))
Global AK_sepchar As String                     ' "separater character" character for AK commands   (default = chr(32))
Global AK_portNumStr As String                  ' Port Number (as a string) for AK commands         (default = "5500")
Global AK_timeout As String                     ' "No AK commands for x seconds" Timeout (as a string)
Global LogAkCommands As Boolean                 ' Log AK commands to the zLog DB file
Global WinsockStateDesc(0 To 9) As String

Global DRIVEPATH As String                              ' drive with root path for all files
Global filepath As String                               ' root path for all files
Global FILEPATH_reports  As String                      ' root path for all report files
Global FILEPATH_data As String                          ' root path for all data files
Global FILEPATH_cal  As String                          ' root path for all calibration files
Global FILEPATH_cfg  As String                          ' root path for all configuration files
Global FILEPATH_log  As String                          ' root path for all log files
Global FILEPATH_rcp  As String                          ' root path for all recipe (and canister recipe) files
Global FILEPATH_manuals  As String                      ' root path for all manuals files
Global FILEPATH_sysdbf  As String                       ' root path for all system database files
Global FILEPATH_backup  As String                       ' default path for backup files

Global DATACAL As String                                ' DataBase File with Calibration info
Global DATAMODEL As String                              ' Template DataBase file
Global DATAMASTER As String                             ' DataBase File with Job Log, Event Log, etc.
Global DATAUSER As String                               ' DataBase File with User Logon info
Global DATARCP As String                                ' DataBase File with Recipe and Canister info
Global DATASYSDEF As String                             ' DataBase File with System info
Global DATAAIRLOG As String                             ' DataBase File with Air Temperature & Humidity log
Global DATAREM As String                                ' DataBase File with REMOTE Tasks
Global DATAZLOG As String                               ' DataBase File with Zlogs (debug logs)

Global CurrTimer As Double                  ' Current System Timer Value
Global DeltTimer As Double                  ' Change in System Timer Value from Previous to Current
Global PrevTimer As Double                  ' Previous System Timer Value
Global currDTS As Date                      ' Current DateTime
Global prevDts As Date                      ' Previous DateTime
Global USINGERRORMSGBYPASS As Boolean       ' Enable Error Handler to "bypass" the Error Message MsgBox (sysdef)
Global ErrorMsgBypassActive As Boolean      ' Allow Error Handler to "bypass" the Error Message MsgBox (this moment)
Global ShortTermErrorCounter As Long        ' Too many errors in as short time >>> override the "bypass"
Global ShortTermErrorMax As Long            ' How many = Too many errors

' MODE, PHASE & CODES DESCRIPTIONS
Global ModeDescLong(0 To MAX_MODE) As String        ' Mode Description (Long version for Station Detail screen)
Global ModeDescShort(0 To MAX_MODE) As String       ' Mode Description (Short version for Main screen)
Global ModeBackColor(0 To MAX_MODE) As String       ' Mode Description Background color
Global ModeForeColor(0 To MAX_MODE) As String       ' Mode Description Foreground (Text) color
Global CycleTypeDesc(0 To 9) As String              ' Cycle Type Description
Global LeakCanisterDesc(0 To 9) As String           ' LeakCheck Canister (Method) Description
Global LeakResultDesc(0 To 9) As String             ' LeakCheck Error Description
Global LeakPhaseDesc(0 To 9) As String              ' LeakCheck Phase Description
Global LoadPhaseDesc(0 To 9) As String              ' Load Phase Description
Global PurgePhaseDesc(0 To 9) As String             ' Purge Phase Description
Global LoadTypeDesc(0 To 9) As String               ' Load Type Description - beginning
Global LoadTypeDesc2(0 To 9) As String              ' Load Type Description - middle
Global LoadTypeDesc3(0 To 9) As String              ' Load Type Description - end
Global StartTypeDesc(0 To 9) As String              ' Start Type Description - beginning
Global StartTypeDesc2(0 To 9) As String             ' Start Type Description - end
Global StartMethodDesc(0 To 9) As String            ' Start Method Description
Global EndMethodDesc(0 To 9) As String              ' End Method Description
Global LeakMethodDesc(0 To 9) As String             ' Leakcheck Method Description
Global LoadMethodDesc(0 To 9) As String             ' Load Method Description
Global PurgeMethodDesc(0 To 9) As String            ' Purge Method Description
Global PurgeTargetDesc(0 To 9) As String            ' Purge Target Description
Global ReportCodeDesc(0 To 19) As String            ' Report Code Description

' PURGE PROFILE STEP-TYPE DESCRIPTIONS
Global PurgeProfileStepDesc(0 To 3) As String       ' Purge Profile Step-Type Description

' Calibration variables
Global STN_Calib_X(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_X2(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_X3(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_X4(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_X5(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_X6(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global Junk(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global STN_Calib_MFC_Range(1 To MAX_STN, 1 To MAX_SHIFT, 0 To MAXMFC) As Single
Global Mfc_Calib_Filename(0 To MAXMFC) As String
Global Mfc_Description(0 To MAXMFC) As String

' Station Process Status Indicators
Global IdlePauseCount As Integer
Global SilenceHornRequest As Boolean
Global CommReadBuffer(0 To MAX_COMM) As String
Global CommReadString(0 To MAX_COMM) As String
Global Port_In_Use(0 To MAX_COMM) As Boolean
Global Port_OK(0 To MAX_COMM) As Boolean
Global Port_Type(0 To MAX_COMM) As String * 1               ' (port) Scale Type
Global Port_Weight(0 To MAX_COMM) As Single                 ' (port) Scale Weight (string converted to number)
Global Port_Value(0 To MAX_COMM) As String                  ' (port) Scale Input String
Global Scale_In_Use(0 To MAX_SCALES) As Boolean
Global Scale_Port(0 To MAX_SCALES) As Integer
Global Scale_OK(0 To MAX_SCALES) As Boolean
Global Scale_Type(0 To MAX_SCALES) As String * 1            ' Scale Type
Global Scale_Weight(0 To MAX_SCALES) As Single              ' Scale Weight (string converted to number)
Global Scale_Value(0 To MAX_SCALES) As String               ' Scale Input String
Global ScalesReadAllOnce As Boolean

' Common Process Variables
Global PAMoisture As Single             ' Current PAS Moisture Value
Global PATemp As Single
Global PAHum As Single
Global AmbMoisture As Single            ' Current Ambient Moisture Value
Global AmbTemp As Single
Global AmbHum As Single
Global AmbBaro As Single
' Leak Check
Global PTinvalue As Single              ' Pressure Transducer value
' LeakTest
Global Deff As Double
Global QN2 As Double
Global Pin As Double
Global Patm As Double
Global SGN2 As Double
Global TN2 As Double
Global DeffCalcFlag As Boolean
Global DeffCalcMsg As String


' just for simulation
Global Sim_ModeStartComplete(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Sim_AuxWt(1 To MAX_STN) As Single                                        ' simulated Aux Canister weight rounded to one decimal
Global Sim_PriWt(1 To MAX_STN) As Single                                        ' simulated Pri Canister weight rounded to one decimal
Global Sim_AuxWt_Current(1 To MAX_STN) As Single                                ' simulated Aux Canister weight
Global Sim_PriWt_Current(1 To MAX_STN) As Single                                ' simulated Pri Canister weight
Global Sim_AuxWt_Last(1 To MAX_STN) As Single
Global Sim_PriWt_Last(1 To MAX_STN) As Single
Global Sim_LastLoadAuxWtChg(1 To MAX_STN) As Single
Global Sim_AuxCan_JobStartPercentLoaded(1 To MAX_STN) As Single
Global Sim_PriCan_JobStartPercentLoaded(1 To MAX_STN) As Single
Global Sim_MfcError(0 To MAX_STN, 0 To MAXMFC) As Single
Global Sim_PasError(1 To 2) As Single
Global Sim_Cycle(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Global Sim_Mode(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Global Sim_LiveFuelDensity As Single
Global Sim_LcPtUser As Integer
Global Sim_LastPTcheckTimer As Double
Global Sim_LeakError(0 To MAX_STN) As Single
Global Purge_Total_Last(1 To MAX_STN, 1 To MAX_SHIFT)

' IO Monitoring Mode Values
Global STN_IOForceMode(0 To MAX_STN) As Integer                                 ' Mode for Io Monitor Screen (0=Common)
Global IOForceActive As Boolean                                                 ' Indicator that An IOForce Screen is being Displayed
Global CalibrateActive As Boolean                                               ' Indicator that the Calibration Screen is being Displayed
Global OptoCommOn_Request As Boolean
Global OptoCommOff_Request As Boolean
Global ReportGenerate_Request As Boolean
Global ReportGenerate_dbPath As String
Global IoComOn_memory As Boolean
Global UsingSimulation_memory As Boolean
Global ScaleCommOn_Request As Boolean
Global ScaleCommOff_Request As Boolean
Global SclComOn_memory As Boolean

' Screen Values
Global Alarm_ForeColor As Long
Global BarActual_ForeColor As Long
Global Data_ForeColor As Long
Global DataBold_ForeColor As Long
Global DataHiLite_ForeColor As Long
Global Entry_ForeColor As Long
Global Good_ForeColor As Long
Global Message_ForeColor As Long
Global Titles_ForeColor As Long
Global TitlesData_Forecolor As Long
Global TitlesLabel_ForeColor As Long
Global Warning_ForeColor As Long
Global Common_BackColor As Long
Global Entry_BackColor As Long
Global EntryInvalid_BackColor As Long
Global EntryUnsaved_BackColor As Long
Global EntryNotChangeable_BackColor As Long
Global StnDetail_ButtonTop As Long
Global OutOfSight As Long
Global SysMessage_Text As String
Global SysMessage_Tooltip As String
Global SysMessage_BackColor As Long
Global SysMessage_ForeColor As Long
Global PurgeAirMsg_Text As String
Global PurgeAirMsg_ToolTip As String
Global PurgeAirMsg_BackColor As Long
Global PurgeAirMsg_ForeColor As Long
Global Stop_In_Progress As Boolean
Global Stn_OperReportNameIsValid As Boolean
Global MasterMode_BackColor As Long
Global StationMode_BackColor As Long

' ********************************************************************
' *************** "RealTime" Variables **************
Global UnreadProgramErrorMessage As Boolean
Global tmrStation As Integer
Global tmrShift As Integer
Global PreviousNow(1 To MAX_STN, 1 To MAX_SHIFT) As Date
Global PreviousReportTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global PreviousTotalTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global DeltaTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global DeltaTotalTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global netmin(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global netflow(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global tarmin(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global tarvol(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global tardone(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean

Global Stn_XYGraph_TestTimer(1 To MAX_STN, 1 To MAX_SHIFT) As Double
Global Stn_XYChart_Xinterval(1 To MAX_STN, 1 To MAX_SHIFT) As Double

' *************** Debug Variables **************
Global Debug_Dts(1 To 19) As Date                                         ' Debug DateTime Stamps
Global Debug_Timer(1 To 19) As Double                                     ' Debug Timer Stamps
Global Debug_Comment(1 To 19) As String                                   ' Debug Comment Message
Global Debug_ZlogPAS_LastUpdate As Double                                 ' Max Allowed Records in the PAS zLog db table
Global Debug_ZlogPAS_MaxRecords As Long                                   ' Max Allowed Records in the PAS zLog db table
Global Debug_ZlogPAS_NumRecords As Long                                   ' Number of Records in the PAS zLog db table
Global Debug_ZlogPAS_Clear As Boolean                                     ' Clear PAS zLog db table flag
Global Debug_ZlogPurge_MaxRecords As Long                                 ' Max Allowed Records in the Purge zLog db table
Global Debug_ZlogPurge_NumRecords As Long                                 ' Number of Records in the Purge zLog db table
Global Debug_ZlogPurge_Clear As Boolean                                   ' Clear Purge zLog db table flag
Global Debug_ZlogScale_MaxRecords As Long                                 ' Max Allowed Records in the Scale zLog db table
Global Debug_ZlogScale_NumRecords As Long                                 ' Number of Records in the Scale zLog db table
Global Debug_ZlogScale_Clear As Boolean                                   ' Clear Scale zLog db table flag

'
' Summary Statistic variables
Type Statistic
    sMin As Single
    sMax As Single
    sAvg As Single
    sCnt As Single
End Type
' Station Statistics variables
Type StnStat
    Btn As Statistic
    Nit As Statistic
    Mix As Statistic
    Pur As Statistic
    AirTemp As Statistic
    AirMoist As Statistic
    FuelTemp As Statistic
End Type
Global StationStatistics(1 To MAX_STN, 1 To MAX_SHIFT) As StnStat


Global WtQueue(1 To 1000, 1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global TimeQueue(1 To 1000, 1 To MAX_STN, 1 To MAX_SHIFT) As Single
Global InIdx(1 To MAX_STN, 1 To MAX_SHIFT) As Integer
Global WTCHGQUEUESIZE As Integer                                        ' 1 to 1000

' The following variables are latching flags used to control station alarm messages
' index starts from zero; index zero is overall for main menu screen
Global Alm_Btn20 As Boolean
Global Alm_Doors As Boolean
Global Alm_Estop As Boolean
Global Alm_ExtContacts As Boolean
Global Alm_Flow As Boolean
Global Alm_LiveFuelLevel(0 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Alm_LiveFuelHeater(0 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Alm_LiveFuelSafety(0 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Alm_LiveFuelN2(0 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Alm_LoadPress(0 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global Alm_N2 As Boolean
Global Alm_UPS As Boolean
Global Alm_Vac(1 To MAX_PRG) As Boolean
Global Alm_Doors_FirstTime As Boolean
Global Alm_Doors_Count As Integer
Global Alm_Doors_Time As Date
Global Alm_SystemVacSw As Boolean
Global Alm_Ups_FirstTime As Boolean
Global Alm_Ups_Count As Integer
Global Alm_Ups_Time As Date
Global Alm_Vac_Count(1 To MAX_PRG) As Integer
'
Global MaintMode As Boolean                     ' true if Maintenance Mode DI is true
'
' General Variables
Global Comm_Flag As Boolean                         ' Communication Status, 1 = good
Global CommErrors(0 To MAX_COMM) As Integer         ' # of comm errors (comm port)
Global CommPrevRead(0 To MAX_COMM) As Double        ' Timer value after last data read (comm port)
Global CurCommPort As Integer                       ' Current MsComm Port being read

Global CurChkStn As Integer             ' Current Station Number for the "Check Stations" Routine
Global DefFunc(2, 9) As Integer         ' Current Function Definition Function #(0=analog;1=digital, 0=common;1=fid;2=station;3=purge)
Global DefPrg As Integer                ' Current PurgeAir Definition PurgeAir Source #
Global DefStn As Integer                ' Current Station Definition Station # (0=Common)
Global DefStnPrg As Integer             ' Current Station Definition PurgeAir #
Global DispStn As Integer               ' Current Detail Stations
Global DispShift As Integer             ' Current shift in use
Global DispCourse As Integer            ' Current Course
Global Maint_Done As Boolean            ' Maintenance done flag
Global Pause_Alarm As Integer           ' NOTPAUSED = 0, SYSTEMPAUSED = 1
Global QboxResponse As Integer
Global Sys_StartTime As Date            ' time when program was last started
Global JobSeqAutoEdit As Boolean

Global systemhasBUTANE As Boolean                       ' At least one station on this system uses Butane
Global systemhasORVR2 As Boolean                        ' At least one station on this system uses Dual-Range MFC's
Global systemhasLIVEFUEL As Boolean                     ' At least one station on this system uses Live Fuel
Global systemhasLEAKTEST As Boolean                     ' At least one station on this system is a LeakTest station
Global systemhasAUTODRAINFILL As Boolean                ' Auto Drain/Fill on this system (i.e. A Fuel Tank with Level Switch, Pump, Valves, etc.)       added 19 Nov2004 MMW
Global systemhasADF_HEATER As Boolean                   ' Yes LiveFuel Tanks have a (optional) Heater
Global systemhasADF_WATERBATH As Boolean                   ' Yes LiveFuel Tanks have a (optional) Heater

' sequence control
Global SEQ_Nmbr(1 To MAX_STN, 1 To MAX_SHIFT) As Integer              ' 1=PostPurge CanVent N2 Feed; 2=unused; 3=unused; 4=unused
Global SEQ_Step(1 To MAX_STN, 1 To MAX_SHIFT) As Integer              ' Current Step Number of Station Sequence.
Global SEQ_StepBeforePause(1 To MAX_STN, 1 To MAX_SHIFT) As Integer   ' Step Number of Station Sequence before Pause.
Global SEQ_Task(1 To MAX_STN, 1 To MAX_SHIFT) As String               ' Description of Current Sequence
Global SEQ_Message(1 To MAX_STN, 1 To MAX_SHIFT) As String            ' Description of Current Step
Global SEQ_Message2(1 To MAX_STN, 1 To MAX_SHIFT) As String            ' Description of Current Step
Global SEQ_EndTime(1 To MAX_STN, 1 To MAX_SHIFT) As Date              ' DateTime When Current Sequence Ended
Global SEQ_Time(1 To MAX_STN, 1 To MAX_SHIFT) As Date                 ' Max DateTime Allowed to Complete Entire Sequence
Global SEQ_Step_Time(1 To MAX_STN, 1 To MAX_SHIFT) As Date            ' Max DateTime Allowed to Complete Current Step
Global SEQ_StartTime(1 To MAX_STN, 1 To MAX_SHIFT) As Date            ' DateTime When Current Sequence Started
Global SEQ_Step_StartTime(1 To MAX_STN, 1 To MAX_SHIFT) As Date       ' DateTime When Current Step Started
Global SEQ_Alarm(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean
Global SEQ_OOT(1 To MAX_STN, 1 To MAX_SHIFT) As Boolean

' Simulation - Live Fuel
Global ADFsim_StorageLevel(1 To MAX_STN) As Single
Global ADFsim_TankLevel(1 To MAX_STN) As Single
Global ADFsim_TankPressure(1 To MAX_STN) As Single
Global ADFsim_TankTemperature(1 To MAX_STN) As Single
Global ADFsim_TankSheathTemp(1 To MAX_STN) As Single
Global ADFsim_PVoverT(1 To MAX_STN) As Single

'
'  Opto
Global Opto_Send_Data(0 To 16) As String    ' Send out to Opto / 16 invalid data flag
Global Opto_Rec_Data(0 To 16) As String     ' Rec in from the Opto / 16 invalid data flag
Global Opto_Error_data As String            ' Current error string from this function
Global Disp_Opto As Integer                 ' Incrementer for opto stations
Global Opto_COMM_ERROR(0 To MAX_ADDR)       ' Current error string from this function
Global Const FULLSCALE = 655350             ' Opto Counts when AI = 10 VDC input
Global OptoReadAllOnce As Boolean

' Temp/Rh Logging
Global CurAirLogFile As String

Global mwDebugInt As Integer

' That's the globals

'******************************************************************************************
'******************************************************************************************
'**********************************************************************************************
Public Declare Function lstrcpy Lib "kernel32" _
   Alias "lstrcpyA" _
   (ByVal lpString1 As String, _
   ByVal lpString2 As String) _
   As Long

Public Declare Function OpenPrinter Lib "winspool.drv" _
   Alias "OpenPrinterA" _
   (ByVal pPrinterName As String, _
   phPrinter As Long, _
   pDefault As PRINTER_DEFAULTS) _
   As Long

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
   (ByVal hPrinter As Long, _
   ByVal level As Long, _
   pPrinter As Byte, _
   ByVal cbBuf As Long, _
   pcbNeeded As Long) _
   As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
   (ByVal hPrinter As Long) _
   As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

' constants for PRINTER_DEFAULTS structure
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ACCESS_ADMINISTER = &H4

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type

Public Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_WARMING_UP = &H10000


Option Explicit

'******************************************************************************************
'******************************************************************************************



Sub Main()
'***********************************************************************
' Function Name:    Main
' Author:           Analytical Process Programmer         8/96
' Description:      This routine is called to start the program.
'                   It loads configuration files and starts the Main Menu
'***********************************************************************
If UseLocalErrorHandler Then On Error GoTo localhandler
SetErrModule 1, 1
Dim msg As String
Dim X As Printer
                
    ' Start Time
    Sys_StartTime = Now()
    UnreadProgramErrorMessage = False
    
    ' Got a Default Printer for us to use?
    PRINTERNAME = "none"
    PRINTERAVAILABLE = False
    ' find the default "Printer" in the collection of "Printers"
    For Each X In Printers
        ' this printer's name
        PRINTERNAME = X.DeviceName
        ' is this printer the default printer?
        If PRINTERNAME = Printer.DeviceName Then
            ' this printer is the Default Printer
            PRINTERAVAILABLE = True
            ' stop looking for any more printers
            Exit For
        End If
        ' this printer isn't the default printer
        PRINTERNAME = "none"
    Next
    
    ' Default File Font (for Detail & Summary Reports)
    DEFAULTFILEFONT.Name = "Lucida Console"     ' Standard windows fixed font
    ' Default Report Font (for other Reports)
    DEFAULTREPORTFONT.Name = "Arial"            ' Standard windows proportional font
    
    ' File Font (for Detail & Summary Reports)
    FILEFONT = DEFAULTFILEFONT                  ' Standard windows fixed font
    FILEFONTSIZE = 10                           ' Size of File Font
    ' Report Font (for other Reports)
    REPORTFONT = DEFAULTREPORTFONT              ' Standard windows proportional font
    
    ' misc constants
    cNO = CInt(0)
    cYES = CInt(1)
    cOFF = CInt(0)
    cON = CInt(1)
    ' (for estimation & simulation only) convert Canister Volume to Canister WorkingCapacity
    DefCanVol2CanWcMult = CSng(14)
    
    ' Revision Levels
    ReportGenRevLvl = RPTGENREVLVL   ' Revision Level Required of Report Generator Program
    CfgRevLvl = CPSREVLVL            ' Current Revision Level for config, sysdef, etc. files
    
    ' set default colors
    SetDefault_ScreenColors
    
    ' set Delays
    DISPDELAY = 10 / (24! * 60! * 60!)      ' 10 sec
    MSGDELAY = 750                          ' 750 msec
    OptoDelay = 0.05    ' seconds
    INTRODELAY = 250    ' milliseconds
    PAUSEDELAY = 950    ' milliseconds
    
    USINGFIDANALYZER = False
    
    IntroDone = False
    VarInitDone = False
    UserLoginOk = False
    ReadyToRun = False
    ShuttingDown = False
    SystemTimers(tmrScanIO).Phase = 0
    
    ' Show Splash Screen
    '   Splash Screen will actually start the cycling of the program
    Load frmAbout
    DoEvents
    frmAbout.txtMessage.text = CStr(Now()) & "   Begin Program Startup" & vbCrLf
    frmAbout.UpdateMsg "   TEST MESSAGE 1" & vbCrLf
    frmAbout.Show
    
    Delay_Box "", INTRODELAY, msgNOSHOW
    frmAbout.txtMessage.text = CStr(Now()) & "   TEST MESSAGE 2" & vbCrLf
    
    ' Check required files & folders
    FilesAndFolders
     frmAbout.txtMessage.text = CStr(Now()) & "   TEST MESSAGE 3" & vbCrLf
           
    ' Setup Station Mode Information
    SetupStnModes
     frmAbout.txtMessage.text = CStr(Now()) & "   TEST MESSAGE 4" & vbCrLf
    
    ' Begin Loading Configuration Files
    frmAbout.UpdateMsg "Load System Definition Files" & vbCrLf
    Delay_Box "", INTRODELAY, msgNOSHOW
    
    ' load screen colors
    Load_ScreenColors
    
    ' Load System Definition Data
    Load_SysDef
        ' Save System Definition Data with updated current program revision
        Save_SysDef
    NR_PRGPROF = MAX_PROFILES
    
    ' Load Node Information
    Load_NodeInfo
    ' Load Opto Information
    Load_OptoInfo
    ' Load Station Information
    Load_StationInfo
    ' Setup Station Information
    SetupStations
    
    ' Setup Function Array Information
    SetupFunctionArrays
    ' Load Analog Function Definition Information
'    If (MAX_ANA_STN > 29) Then Update_AnalogFuncDef 29, MAX_ANA_STN
    Load_AnalogFuncDef
    ' Load Digital Function Definition Information
    Load_DigitalFuncDef
    
    ' Load Configuration File
    Load_Config
    
    frmAbout.UpdateMsg "Module 1 ln3947 Load Config" & vbCrLf
    
    ' Load PurgeAir Source Definition Information
    Load_PurgeInfo
    frmAbout.UpdateMsg "Module 1 ln3951  Load PurgeInfo" & vbCrLf
    
    ' Load Scale Configuration File
    Load_ScaleConfig
    frmAbout.UpdateMsg "Module 1 ln3955  Load ScaleConfig" & vbCrLf
    
    ' Load LeakTest Cfg & Rcp File
    Load_LeakTest
    frmAbout.UpdateMsg "Module 1 ln3959  Load LeakTest" & vbCrLf
    
    
    ' Check Cfg Revision Level, upgrade if required
    '     CpsRevLvl is the desired rev level of the program
    '     CfgRevLvl is the current rev level of the config & sysdef files on this system
    If CfgRevLvl = CPSREVLVL Then
    
        ' ********************************************************
        ' Read Recipe, Canister and Simulation Data from Databases
        ' ********************************************************
        ' Load Station Canister Information
        Load_StationCanisters
        ' Load Station Recipes
        Load_StationRecipes
        ' Load Station Sequences
        Load_StationSequences
        ' Load Station PurgeProfiles
        Load_StationProfiles
        ' Load Butane Supply Data
        Load_ButaneSupply
        ' Load Station Config Data
        Load_StationConfig
        ' Load Station Simulation Data
        Load_Simulation
        ' Load Station ADF Config Data
        Load_AdfConfig
        ' Load Controllers Config Data
        Load_Controllers
        ' Load Local PAS Control Config Data
        Load_LocalPAS
        frmAbout.UpdateMsg "Module 1 ln3990  Load Local PAS" & vbCrLf
        ' Load FID Config Data
        ' Load_FidConfig
        ' ***************************************************************************
        ' Copy Master Canisters, Recipes, PurgeProfiles and JobSequences to Remote DB
        ' ***************************************************************************
        DispStn = 1
        ' Master Canisters, Recipes, PurgeProfiles & JobSequences
        UpdateAllRemoteMasters
        frmAbout.UpdateMsg "Module 1 ln3999  update all remote masters" & vbCrLf
        
        'frmCanRecipe.Hide
        'frmCanRecipe.
        'Unload frmCanRecipe
        'frmAbout.UpdateMsg "Updated Remote Master Canisters" & vbCrLf
        ' Master Recipes
        'frmRecipe.Hide
        'frmRecipe.CopyAllRemoteMasterRecipes
        'Unload frmRecipe
        'frmAbout.UpdateMsg "Updated Remote Master Recipes" & vbCrLf
        ' Master PurgeProfiles
        'frmPurgeProfile.Hide
        'frmPurgeProfile.CopyAllRemoteMasterProfiles
        'Unload frmPurgeProfile
        'frmAbout.UpdateMsg "Updated Remote Master PurgeProfiles" & vbCrLf

    Else
        msg = "The Configuration files are too old for this version !!!" & vbCrLf & vbCrLf
        msg = msg & "Cannot continue. Please fix the problem then try again." & vbCrLf
        MsgBox msg, vbCritical, "SHUTTING DOWN!"
        End
    End If
    
    
    ' Setup Opto Information
    frmAbout.UpdateMsg "Setup Opto Information" & vbCrLf
    SetupOpto
    Delay_Box "", INTRODELAY, msgNOSHOW
    
    ' Setup Scale/Port Information
    frmAbout.UpdateMsg "Setup Scale/Port Information" & vbCrLf
    frmComm8Card.Setup_Scales
    Delay_Box "", INTRODELAY, msgNOSHOW
    
    ' Write to Event Log, starting system
    frmAbout.UpdateMsg "CPS Reporting System Restarted" & vbCrLf
    Write_ELog "CPS Reporting System Restarted"
    Delay_Box "", INTRODELAY, msgNOSHOW
    ' Set Password Defaults
    Init_Password
    If AutoLogon <> 0 Then Delay_Box "", INTRODELAY, msgNOSHOW
    If SysConfig.AutoLogon <> 0 Then Delay_Box "", INTRODELAY, msgNOSHOW
    ' Initialize Variables
    Init_Vars
    ' Open Temp/Rh Log File
    If LogTempRh Then
        If AirLogFileIsReady(Now) Then
            frmAbout.UpdateMsg "Current AirLog database file is " & CurAirLogFile & vbCrLf
            Write_ELog ("Current AirLog database file is " & CurAirLogFile)
        Else
            frmAbout.UpdateMsg "AirLog database file not available" & vbCrLf
            Write_ELog ("AirLog database file not available")
        End If
    End If
    ' Clear(Reset) any Active REMOTE Tasks
    If USINGREMCANLOAD Then
        RemTask_ClearActive
    End If
            
    ' System Init Complete
    About_Counter = 0
    VarInitDone = True
    UserLoginOk = True
    ReadyToRun = True
    frmAbout.cmdSetup.Visible = True
    '   note:   user will have 5 seconds to pause system start
    '           by selecting System Definition on the About Screen.
    '           The startup sequence will resume after
    '           all SystemDefinition screens are closed (password screen too).
    frmAbout.UpdateMsg "System Initialization Complete" & vbCrLf
    
    ' ***************************************************************
    '   Splash Screen (frmAbout) will start the cycling of the program
    '   Main Process Loop is on frmMainMenu
    ' ***************************************************************
    
    
ResetErrModule
Exit Sub

localhandler:
Dim iresponse As Integer
iresponse = ErrorHandler(err)
Select Case iresponse
  Case vbAbort       ' Exit if abort
    ResetErrModule
    frmMainMenu.MousePointer = vbDefault
    frmDelayBox.MousePointer = vbDefault
    Exit Sub
  Case vbRetry       ' try error line again
    Resume
  Case vbIgnore      ' Skip to next line, try to ignore
    Resume Next
End Select
End Sub

Public Sub menuExit()
Dim sString As String
    ' Exit Program
    If CheckPass("G", True) Then
        sString = "You are about to Exit the Program!" & vbCrLf & _
            vbCrLf & "Exiting will stop all currently running stations." & vbCrLf _
            & vbCrLf & "Please confirm Exit; or Cancel"
        frmCheckIt.CheckIt 1, sString
    End If
End Sub

Public Sub menuLogin()
    If CheckPass("J", True) Then
        frmPassword.Show
    End If
End Sub

Public Sub menuLogout()
    If CheckPass("J", True) Then
        Logout
    End If
End Sub

Public Sub menuPrintFile()
    If CheckPass("F", True) Then
        frmPrint.Show
    End If
End Sub

Public Sub menuCopyFile()
    If CheckPass("F", True) Then
        frmCopy.Show
    End If
End Sub

Public Sub menuConfiguration()
    If CheckPass("O", True) Then
        frmConfig.Show
    End If
End Sub

Public Sub menuCalibration()
    If CheckPass("X", True) Then
        frmSelCal.Show
    End If
End Sub

Public Sub menuSysdef()
    ' Select System Definition
    If CheckPass("H", True) Then
        frmSysDefMain.Show
    End If
End Sub

Public Sub menuIoMonitor()
    If CheckPass("2", True) Then
        frmIoMonitor.Show
    End If
End Sub

Public Sub menuOperatorManual()
    ShowDoc "OperatorManual"
End Sub

Public Sub menuEventLog()
    If CheckPass("Z", True) Then
        frmEventLog.Show
    End If
End Sub

Public Sub menuFirstAid()
    If CheckPass("T", True) Then
        frmFirstAid.Show
    End If
End Sub

Public Sub menuJobList()
    If CheckPass("C", True) Then
        If CheckPass("M", True) Then
            frmJoblist.Show
        End If
    End If
End Sub

Public Sub menuStnDetail()
'    DispStn = 1
'    DispShift = 1
    frmStnDetail.Show
End Sub

Public Sub menuAbout()
    'About
    If CheckPass("D", True) Then
        frmAbout.Show
    End If
End Sub

Public Sub menuAk_Client()
    'AK Client
    If CheckPass("7", True) Then
        frmAkClient.Show
    End If
End Sub

Public Sub menuAk_Server()
    'AK Server
    If CheckPass("7", True) Then
        frmAKServer.Show
    End If
End Sub

Public Sub menuOotMonitor()
    frmOOTmonitor.Show
End Sub

Public Sub menuOverview()
    frmMainMenu.Show
End Sub

Public Sub menuReview()
    ' Review Previous Cycle Data
    If CheckPass("F", True) Then
        frmReview.Show
    End If
End Sub

Public Sub menuWatch()
    ' Watch Current Cycle Data
    If CheckPass("F", True) Then
        frmDataWatcher.Show
    End If
End Sub

Public Sub menuCanisters()
    If CheckPass("N", True) Then
        frmCanRecipe.Show
        frmCanRecipe.ChgCanRcpMode (CInt(MASTERMODE))
        frmCanRecipe.Hide
        ' open Master Canisters selection screen
        frmSearchCan.Show
    End If
End Sub

Public Sub menuRecipes()
    If CheckPass("N", True) Then
        frmRecipe.Show
        frmRecipe.ChgRecipeMode (CInt(MASTERMODE))
        frmRecipe.Hide
        ' open Master Recipes selection screen
        frmSearchRcp.Show
        frmSearchRcp.ChgSelectionDestination rcpdestRecipe
    End If
End Sub

Public Sub menuCourses()
    If CheckPass("N", True) Then
        frmCourses.Show
        frmCourses.ChgJobSeqMode (CInt(MASTERMODE))
        frmCourses.Hide
        ' open Master Job Sequences selection screen
        frmSearchJobSeq.Show
    End If
End Sub

Public Sub menuPurgeProfiles()
    If CheckPass("N", True) Then
        frmPurgeProfile.Show
        frmPurgeProfile.ChgProfileMode (CInt(MASTERMODE))
        frmPurgeProfile.Hide
        ' open Master Profiles selection screen
        frmSearchProf.Show
        frmSearchProf.ChgSelectionDestination CInt(profdestProfile)
    End If
End Sub

Public Sub menuRemCanLoad()
    If CheckPass("N", True) Then
        frmSearchRemote.Show
    End If
End Sub

Public Sub menuScaleMonitor()
    If CheckPass("3", True) Then
        frmComm8Card.Show
    End If
End Sub

Public Sub menuLeakTest()
Dim cntr As Integer
    If CheckPass("K", True) Then
        cntr = 0
        Do Until ((STN_INFO(DispStn).Type = STN_LEAKTEST_TYPE) Or (cntr > NR_STN))
            If (STN_INFO(DispStn).Type <> STN_LEAKTEST_TYPE) Then
                DispStn = IIf((DispStn < NR_STN), (DispStn + 1), 1)
            End If
            cntr = cntr + 1
        Loop
        frmLeakTest.Show
    End If
End Sub

Public Sub menuViewAirLog()
    If LogTempRh Then
        frmViewAirLog.Show
    Else
        Delay_Box "Air Log not in use on this system", MSGDELAY, msgSHOW
    End If
End Sub

Public Sub menuViewFuelUseLog()
    frmViewFuelUseLog.Show
End Sub

Public Sub menuBeaconOff()
    Com_OutDigital icAlarmBeacon, cOFF
End Sub

Public Sub menuHornOff()
    SilenceHornRequest = True
End Sub

Public Sub menuButane()
    If systemhasBUTANE Then
        FrmButaneVolume.Show
    Else
        Delay_Box "No Butane on this system", MSGDELAY, msgSHOW
    End If
End Sub

Public Sub menuSimulation()
    ' Simulation
    If Not IoComOn And USINGSIMULATION Then
        frmSimCntrlPnl.Show
    End If
End Sub

Public Sub ShowDoc(ByVal sManualName As String)
' Uses bat file to display designated manual
'   if pdf version of the manual exists it is opened
'   otherwise the doc or rtf version (if it exists) is opened
Dim sString As String
Dim fs As Object
Set fs = CreateObject("Scripting.FileSystemObject")
    sString = ""
    ' Which version(s) of the manual exist?   *.pdf? *.doc? *.rtf?
    If fs.FileExists(FILEPATH_manuals & sManualName & ".pdf") Then
        ' pdf version of the manual exists; use it
        sString = FILEPATH_manuals & "showDoc.bat  " & sManualName & ".pdf"
    ElseIf fs.FileExists(FILEPATH_manuals & sManualName & ".doc") Then
        ' doc version of the manual exists; use it
        sString = FILEPATH_manuals & "showDoc.bat  " & sManualName & ".doc"
    ElseIf fs.FileExists(FILEPATH_manuals & sManualName & ".rtf") Then
        ' rtf version of the manual exists; use it
        sString = FILEPATH_manuals & "showDoc.bat  " & sManualName & ".rtf"
    End If
    If sString <> "" Then
        ' Run bat to open documentation
        Shell sString
    Else
        ' no version of the manual exists
        sString = "The " & sManualName & " manual cannot be located." & vbCrLf
        AckMsg_Box sString
    End If
End Sub

