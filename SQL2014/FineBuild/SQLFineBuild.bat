@ECHO OFF
REM SQL FineBuild   
REM Copyright FineBuild Team � 2008 - 2021.  Distributed under Ms-Pl License
REM
REM Created 30 Jun 2008 by Ed Vassie V1.0 

REM Setup Script Variables
SET SQLCRASHID=
IF '%SQLDEBUG~0,1%' NEQ '/' SET SQLDEBUG=
IF '%SQLFBDEBUG%' == '' SET SQLFBDEBUG=REM
SET SQLFBCMD=%~f0
SET SQLFBPARM=%*
SET SQLFBPARM=%SQLFBPARM:           /= /%
SET SQLFBPARM=%SQLFBPARM:          /= /%
SET SQLFBPARM=%SQLFBPARM:         /= /%
SET SQLFBPARM=%SQLFBPARM:        /= /%
SET SQLFBPARM=%SQLFBPARM:       /= /%
SET SQLFBPARM=%SQLFBPARM:      /= /%
SET SQLFBPARM=%SQLFBPARM:     /= /%
SET SQLFBPARM=%SQLFBPARM:    /= /%
SET SQLFBPARM=%SQLFBPARM:   /= /%
SET SQLFBPARM=%SQLFBPARM:  /= /%
SET SQLFBFOLDER=%~dp0
FOR /F "usebackq tokens=*" %%X IN (`CHDIR`) DO (SET SQLFBSTART=%%X)
SET SQLLOGTXT=
SET SQLRC=0
SET SQLPROCESSID=
SET SQLTYPE=
SET SQLUSERVBS=
IF '%SQLVERSION%' == '' SET SQLVERSION=SQL2014
CALL "%SQLFBFOLDER%\Build Scripts\Set-FBVersion"
%WINDIR%\SYSTEM32\REGSVR32 /s VBSCRIPT.DLL

%SQLFBDEBUG% %TIME:~0,8% Validate Parameters

ECHO '?' '/?' '-?' 'HELP' '/HELP' '-HELP' | FIND /I "'%1'" > NUL
IF %ERRORLEVEL% == 0 GOTO :HELP

GOTO :RUN

:RUN
%SQLFBDEBUG% %TIME:~0,8% Run the install
ECHO.
ECHO SQL FineBuild %SQLFBVERSION% for %SQLVERSION%
ECHO Copyright FineBuild Team (c) 2008 - %DATE:~6,4%.  Distributed under Ms-Pl License
ECHO SQL FineBuild Wiki: https://github.com/SQL-FineBuild/Common/wiki
ECHO Run on %COMPUTERNAME% by %USERNAME% at %TIME:~0,8% on %DATE%:
ECHO %0 %SQLFBPARM%
ECHO.
ECHO ************************************************************
ECHO %TIME:~0,8% *********** FineBuild Configuration starting

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Logfile (Prepare Log File)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:LogFile %SQLFBPARM%`) DO (SET SQLLOGTXT=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process LogFile var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Debug (Check Debugging requirements)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Debug %SQLFBPARM%`) DO (SET SQLDEBUG=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process Debug var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
IF '%SQLDEBUG%' NEQ '' SET SQLFBDEBUG=ECHO

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  FBParm (Refresh %SQLFBPARM)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:FBParm %SQLFBPARM%`) DO (SET SQLFBPARM=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process FBParm var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  FBParm (Minify %SQLFBPARM)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:MiniParm %SQLFBPARM%`) DO (SET SQLFBPARM=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process MiniParm var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  ProcessId (Refresh %SQLPROCESSID)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:ProcessId %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLPROCESSID=%%X)
SET SQLRC=%ERRORLEVEL%
%SQLFBDEBUG% ProcessId label: '%SQLPROCESSID%'
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process ProcessId var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Type (Refresh %SQLTYPE)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Type %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLTYPE=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process Type var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Build FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigBuild.vbs" %SQLFBPARM% %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

%SQLFBDEBUG% %TIME:~0,8% Report FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigReport.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

IF '%SQLPROCESSID%' GTR 'R2' GOTO :Refresh
IF '%SQLPROCESSID%' NEQ '' GOTO :%SQLPROCESSID%

:R1
ECHO %TIME:~0,8% *********** %SQLVERSION% Preparation processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild1Preparation.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

ECHO %TIME:~0,8% *********** Refreshing environment variables

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Temp (Refresh %TEMP)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Temp %SQLFBPARM% %SQLDEBUG%`) DO (SET TEMP=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process TEMP var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
SET TMP=%TEMP%

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  PathPS (Refresh %PSMODULEPATH)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:PathPS %SQLFBPARM% %SQLDEBUG%`) DO (SET PSModulePath=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process PSMODULEPATH var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

:R2
ECHO %TIME:~0,8% *********** %SQLVERSION% Install processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild2InstallSQL.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

:Refresh

ECHO %TIME:~0,8% *********** Refreshing environment variables

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Temp (Refresh %TEMP)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Temp %SQLFBPARM% %SQLDEBUG%`) DO (SET TEMP=%%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process TEMP var failed
IF %SQLRC% NEQ 0 GOTO :ERROR
SET TMP=%TEMP%

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  Path (Refresh %PATH)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:Path %SQLFBPARM% %SQLDEBUG%`) DO (PATH %%X)
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% == 1 SET SQLRC=0
IF %SQLRC% NEQ 0 ECHO Process PATH var failed
IF %SQLRC% NEQ 0 GOTO :ERROR

IF '%SQLPROCESSID%' GTR 'R2' GOTO :%SQLPROCESSID%

:R3
ECHO %TIME:~0,8% *********** %SQLVERSION% Fixes processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild3InstallFixes.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR
IF '%SQLTYPE%' == 'FIX' GOTO :COMPLETE

:R4
ECHO %TIME:~0,8% *********** %SQLVERSION% Xtras processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild4InstallXtras.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

:R5
ECHO %TIME:~0,8% *********** %SQLVERSION% Configuration processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild5ConfigureSQL.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

:R6
ECHO %TIME:~0,8% *********** User Setup processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FineBuild6ConfigureUsers.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

:COMPLETE
IF EXIST "%TEMP%\FBCMDRUN.BAT" DEL /F "%TEMP%\FBCMDRUN.BAT"
ECHO.
ECHO ************************************************************
ECHO *  
ECHO * %SQLVERSION% FineBuild Install Complete  
ECHO *
ECHO ************************************************************

GOTO :END

:RD
ECHO %TIME:~0,8%             SQL Configuration Discovery processing
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigDiscover.vbs" %SQLDEBUG%
SET SQLRC=%ERRORLEVEL%
IF %SQLRC% NEQ 0 GOTO :ERROR

GOTO :END

:ERROR

IF %SQLRC% == 3010 GOTO :REBOOT

%SQLFBDEBUG% %TIME:~0,8% ConfigVar: CrashId (Refresh %SQLCRASHID)
FOR /F "usebackq tokens=*" %%X IN (`CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:CrashId %SQLFBPARM% %SQLDEBUG%`) DO (SET SQLCRASHID=%%X)
ECHO %TIME:~0,8% Stopped in Process Id %SQLCRASHID%
ECHO.
%SQLFBDEBUG% %TIME:~0,8% ConfigVar: LogView (Display FineBuild Log File)
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:LogView
ECHO %TIME:~0,8% Bypassing remaining processes

GOTO :END

:REBOOT

ECHO.
ECHO ************************************************************
ECHO *  
ECHO * %SQLVERSION% FineBuild  ********** REBOOT IN PROGRESS **********  
ECHO *
ECHO ************************************************************

GOTO :END

:HELP

ECHO Usage: %0 [/Type:Fix/Full/Client/Workstation] [...]
ECHO.
ECHO SQLFineBuild.bat accepts a large number of parameters.  See Fine Install Options in the FineBuild Wiki for details.
ECHO.

SET SQLRC=4
GOTO :EXIT

:R7
REM End point for most processing
:R8
REM End Point for /ReportOnly:Yes
:END

%SQLFBDEBUG% %TIME:~0,8% Report FineBuild Configuration
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigReport.vbs" %SQLDEBUG%

%SQLFBDEBUG% %TIME:~0,8% ConfigVar:  ReportView (Display FineBuild Configuration Report)
CSCRIPT //nologo "%SQLFBFOLDER%\Build Scripts\FBConfigVar.vbs" /VarName:ReportView

ECHO.
ECHO ************************************************************
ECHO *                                           
ECHO * %0 process completed with code %SQLRC%   
ECHO *
ECHO * Log file in %SQLLOGTXT%
ECHO *                                           
ECHO ************************************************************

GOTO :EXIT

:R9
:EXIT
EXIT /B %SQLRC%
