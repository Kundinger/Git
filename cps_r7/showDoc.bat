@ECHO OFF
SETLOCAL 
REM  Bat file to open a CPS_r7 Document 
REM  (the document name is passed to this bat file as command line argument #1)  
REM      1st, make sure that there is something to do
   IF "%1%" EQU "" GOTO NOWORK
   @ECHO The filename "%1%" was passed to this bat file.
   GOTO ENDWORK
   :NOWORK
   @ECHO.
   @ECHO NO filename was passed to this bat file. Nothing to do; Exiting!
   @ECHO.
   PAUSE
   GOTO END
   :ENDWORK
REM      2nd, find the right folder
   IF "%WHEREISCPS%" EQU "" GOTO NOPATH
   @ECHO The WHEREISCPS environment variable is set to "%WHEREISCPS%".
   SET MANUALSPATH=%WHEREISCPS%\manuals\
   GOTO ENDPATH
   :NOPATH
   @ECHO The WHEREISCPS environment variable was NOT detected; using default folder c:\cps_r7.
   SET MANUALSPATH=c:\cps_r7\manuals\
   :ENDPATH
REM     3rd, find the right file
   IF NOT EXIST "%MANUALSPATH%"%1% GOTO NOFILE
   @ECHO The "%1%" file was found 
   @ECHO  in "%MANUALSPATH%". Opening it!
   START  %MANUALSPATH%%1%
   GOTO ENDFILE
   :NOFILE
   @ECHO.
   @ECHO The  "%1%" file was NOT found 
   @ECHO  in "%MANUALSPATH%"; Exiting!
   @ECHO.
   PAUSE
   GOTO END
   :ENDFILE
:END
rem opause
ENDLOCAL
@ECHO ON
