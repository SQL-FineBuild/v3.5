@ECHO OFF
REM Copyright FineBuild Team � 2018.  Distributed under Ms-Pl License
REM
CALL "SQLFineBuild.bat" %*     /Type:Full                    /IAcceptLicenseTerms             ^
 /SAPWD:"UseAL0ngPa55phrase!"  /GroupDBA:"GBGGDBAS01"        /GroupDBANonSA:"GBGGDBAN01"      ^
 /SetupSQLDB:YES                             ^
 /SetupSQLASCluster:YES                      ^
 /SetupSQLRSCluster:YES                      ^
 /SetupSQLIS:YES                             ^
 /SetupAlwaysOn:YES                          ^
 /SetupAODB:YES                              ^
 /SQLSVCAccount:"ROOT\ServGB_SQLDB_1$"       ^
 /AGTSVCACCOUNT:"ROOT\ServGB_SQLAG_1$"       ^
 /ASSVCACCOUNT:"ROOT\ServGB_SQLAS_1$"        ^
 /FTSVCACCOUNT:"ROOT\ServGB_SQLFT_1$"        ^
 /ISSVCACCOUNT:"ROOT\ServGB_SQLIS_1$"        ^
 /RSSVCACCOUNT:"ROOT\ServGB_SQLRS_1$"        ^
 /BROWSERSVCACCOUNT:"ROOT\ServGB_SQLBR_1$"   ^
 /VolProg:C /VolTempWin:C                    ^
 /VolData:J /VolLog:K /VolTemp:T             ^
 /VolBackup:I /VolDataFS:K /VolDataFT:K      ^
 /VolDataAS:F /VolLogAS:G /VolTempAS:F       ^
 /VolBackupAS:G                              ^
 /SetupCmdShell:Yes                          ^
 /CmdshellAccount:"ROOT\AppGB_SQLCS_0001"    /CmdshellPassword:"j25Fb*ef$36ySIyBW7hZ"         ^
 /SetupRSExec:Yes                            ^
 /RSEXECACCOUNT:"ROOT\APPGB_SQLRS_0001"      /RSEXECPASSWORD:"Prf53g#fdf$Efbv8QGH3"