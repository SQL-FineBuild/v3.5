@ECHO OFF
REM Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
REM
REM /Type:Rebuild allows re-running a portion of SQL FineBuild after the initial build has completed successfully
REM The example in this script refreshes the Generic Maintenance procedures
REM
CALL "SQLFineBuild.bat" %*     /Type:Rebuild                 /IAcceptLicenseTerms          ^
 /SetupGenMaint:Yes                                                                        ^
 /Restart:5EF