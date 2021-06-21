@ECHO OFF
REM Copyright FineBuild Team © 2021.  Distributed under Ms-Pl License
REM
REM /Type:Refresh allows re-running a portion of SQL FineBuild after the initial build has completed successfully
REM The example in this script refreshes the Generic Maintenance procedures
REM
CALL "SQLFineBuild.bat" %*     /Type:Refresh                 /IAcceptLicenseTerms          ^
 /SetupGenMaint:Yes                                                                        ^
 /Restart:5EF