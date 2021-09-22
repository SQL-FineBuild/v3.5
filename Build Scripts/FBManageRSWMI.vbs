'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageRSWMI.vbs  
'  Copyright FineBuild Team © 2018 - 2021.  Distributed under Ms-Pl License
'
'  Purpose:      Manage RS WMI processes 
'
'  Author:       Ed Vassie
'
'  Date:         26 Oct 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     26 Oct 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageRSWMI: Set FBManageRSWMI = New FBManageRSWMIClass

Dim objRSInParam, objRSOutParam
Dim intRSLcid

Class FBManageRSWMIClass
  Dim objRSConfig, objRSWMI, objShell
  Dim strFunction, strHTTP, strInstRS, strInstRSSQL, strOSVersion, strPath, strRSAlias, strRSNamespace, strRSNetName, strRSWMIPath
  Dim strServer, strSetupPowerBI, strSetupSSL, strSetupSQLRSCluster, strSSLCertThumb, strTCPPortRS, strTCPPortSSL, strSQLVersion, strWMIPath


Private Sub Class_Initialize
  Call DebugLog("FBManageRSWMI Class_Initialize:")
  Dim strInstRSWMI, strRSVersionNum

  Set objShell      = WScript.CreateObject ("Wscript.Shell")

  strHTTP           = GetBuildfileValue("HTTP")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstRSSQL      = GetBuildfileValue("InstRSSQL")
  strInstRSWMI      = GetBuildfileValue("InstRSWMI")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strRSAlias        = GetBuildfileValue("RSAlias")
  strRSNamespace    = "MSReportServer_ConfigurationSetting"
  strRSNetName      = GetBuildfileValue("RSNetName")
  strRSVersionNum   = GetBuildfileValue("RSVersionNum")
  strServer         = GetBuildfileValue("AuditServer")
  strSetupPowerBI   = GetBuildfileValue("SetupPowerBI")
  strSetupSSL       = GetBuildfileValue("SetupSSL")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strSSLCertThumb   = GetBuildfileValue("SSLCertThumb")
  strTCPPortRS      = GetBuildfileValue("TCPPortRS")
  strTCPPortSSL     = GetBuildfileValue("TCPPortSSL")

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\v" & strRSVersionNum & "\Admin"
    Case strSQLVersion >= "SQL2017"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\V" & strRSVersionNum & "\Admin"
    Case strSetupPowerBI = "YES"
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\V" & strRSVersionNum & "\Admin"
    Case Else
      strRSWMIPath  = "winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ReportServer\" & strInstRSWMI & "\v" & strRSVersionNum & "\Admin"
  End Select

End Sub


Function RunRSWMI(strFunction, strOK)
  Call DebugLog("RunRSWMI: " & strFunction)
' MSReportManager_ConfigurationSetting: https://docs.microsoft.com/en-us/previous-versions/sql/sql-server-2005/ms154684(v=sql.90)
' MSReportServer_ConfigurationSetting:  https://docs.microsoft.com/en-us/previous-versions/sql/sql-server-2005/ms154648(v=sql.90)
  Dim intErrSave

  strWMIPath        = strRSNamespace & ".InstanceName='" & strInstRSSQL & "'"
  strDebugMsg1      = "WMI Path: " & strWMIPath
  Set objRSOutParam = objRSWMI.ExecMethod(strWMIPath, strFunction, objRSInParam)
  intErrSave        = objRSOutParam.HRESULT
  If intErrSave = -2147023181 Then
    WScript.Sleep GetBuildfileValue("WaitLong")
    WScript.Sleep GetBuildfileValue("WaitLong")
    Set objRSOutParam = objRSWMI.ExecMethod(strWMIPath, strFunction, objRSInParam)
    intErrSave      = objRSOutParam.HRESULT
  End If
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      intErrSave    = 0
    Case intErrSave = -2147023181
      Call SetBuildMessage(strMsgError,   "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction & ": " & "RPC Server not available")
    Case strSQLVersion = "SQL2005"
      Call SetBuildMessage(strMsgWarning, "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction)
    Case Else
      Call SetBuildMessage(strMsgWarning, "Unexpected WMI error: " & CStr(intErrSave) & " for " & strFunction & ": " & objRSOutParam.Error)
  End Select

  RunRSWMI        =  intErrSave

End Function


Function SetRSInParam(strFunction)
  Call DebugLog("SetRSInParam: " & strFunction)

  If Not IsObject(objRSWMI) Then
    Call SetRSWMI()
  End If

  strWMIPath        = strRSNamespace & ".InstanceName='" & strInstRSSQL & "'"
  strDebugMsg1      = "WMI Path: " & strWMIPath
  Set objRSConfig   = objRSWMI.Get(strWMIPath)
  Set objRSInParam  = objRSConfig.Methods_(strFunction).inParameters.SpawnInstance_()

  SetRSInParam      = strFunction

End Function


Sub SetRSDatabase(strServer, strRSDBName)
  Call DebugLog("SetRSDatabase: " & strRSDBName & " on "  & strServer)

  strFunction       = SetRSInParam("SetDatabaseConnection")
  objRSInParam.Properties_.Item("Server")            = strServer
  objRSInParam.Properties_.Item("DatabaseName")      = strRSDBName
  objRSInParam.Properties_.Item("CredentialsType")   = 2 ' Use Windows Service account
  objRSInParam.Properties_.Item("UserName")          = ""
  objRSInParam.Properties_.Item("Password")          = ""
  Call RunRSWMI(strFunction, "")

End Sub


Sub SetRSDirectory(strApplication, strDirectory)
  Call DebugLog("SetRSDirectory: " & strApplication & ", " & strDirectory)
  Dim strStoreNamespace

  strStoreNamespace = strRSNamespace

  Select Case True
    Case strSQLVersion > "SQL2005"
      strFunction   = SetRSInParam("SetVirtualDirectory")
    Case strApplication = "ReportManager"
      strRSNamespace = "MSReportManager_ConfigurationSetting"
      strFunction    = SetRSInParam("CreateVirtualDirectory")
    Case Else
      strFunction   = SetRSInParam("CreateVirtualDirectory")
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      objRSInParam.Properties_.Item("IISPath")          = "IIS://localhost/w3svc/1/root"
      objRSInParam.Properties_.Item("Name")             = strDirectory
    Case strSetupPowerBI = "YES"
      objRSInParam.Properties_.Item("Application")      = strApplication
      objRSInParam.Properties_.Item("VirtualDirectory") = strDirectory
      objRSInParam.Properties_.Item("Lcid")             = intRSLcid
    Case Else
      objRSInParam.Properties_.Item("Application")      = strApplication
      objRSInParam.Properties_.Item("VirtualDirectory") = strDirectory
      objRSInParam.Properties_.Item("lcid")             = intRSLcid
  End Select
  Call RunRSWMI(strFunction, "-2147220930 -2147220978 -2147220980 -2147024864") ' OK if Directory already exists

  Select Case True
    Case strSQLVersion > "SQL2005"
      Call SetRSURL(strApplication, strDirectory)
    Case strApplication = "ReportManager"
      strFunction   = SetRSInParam("SetReportManagerIdentity")
      objRSInParam.Properties_.Item("ApplicationPool")  = "DefaultAppPool"
      Call RunRSWMI(strFunction, "")
    Case Else
      strFunction   = SetRSInParam("SetWebServiceIdentity")
      objRSInParam.Properties_.Item("ApplicationPool")  = "DefaultAppPool"
      Call RunRSWMI(strFunction, "")
  End Select

  strRSNamespace    = strStoreNamespace

End Sub


Private Sub SetRSURL(strApplication, strDirectory)
  Call DebugLog("SetRSURL: " & strApplication)
  Dim strURLVar

  Select Case True
    Case strSetupPowerBI = "YES"
      strFunction    = SetRSInParam("ReserveURL")
      objRSInParam.Properties_.Item("Application")                  = strApplication
      objRSInParam.Properties_.Item("Lcid")                         = intRSLcid
      strURLVar      = "UrlString"
    Case Else
      strFunction    = SetRSInParam("ReserveURL")
      objRSInParam.Properties_.Item("Application")                  = strApplication
      objRSInParam.Properties_.Item("lcid")                         = intRSLcid
      strURLVar      = "URLString"
  End Select

  Call SetRSURLItem(strFunction, objRSInParam, strURLVar, strServer, strApplication)
  If strRSNetName <> strServer Then
    Call SetRSURLItem(strFunction, objRSInParam, strURLVar, strRSNetName, strApplication)
  End If

End Sub


Private Sub SetRSURLItem(strFunction, objRSInParam, strURLVar, strHost, strApplication)
  Call DebugLog("SetRSURLItem: " & strHost)
  Dim strURL

  Select Case True
    Case strHost = strServer
      strURL         = "HTTP://" & UCase(strHost) & ":" & strTCPPortRS
    Case strSetupSSL = "YES"
      strURL         = "HTTPS://+:" & strTCPPortSSL
      Call SetRSSSLBind(strApplication)
    Case Else
      strURL         = "HTTP://" & UCase(strHost) & ":" & strTCPPortRS
  End Select

  strDebugMsg1       = "URL: " & strURL
  objRSInParam.Properties_.Item(CStr(strURLVar)) = strURL
  Call RunRSWMI(strFunction, "-2147220932") ' OK if URL already exists

End Sub


Sub SetRSSSL()
  Call DebugLog("SetRSSSL:")
  Dim strStoreNamespace

  strStoreNamespace = strRSNamespace

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strRSNamespace = "MSReportManager_ConfigurationSetting"
      Call SetRSSSLBind("ReportServerWebService")  
      Call SetRSSSLBind("ReportManager") 
    Case strSetupPowerBI = "YES"
      Call SetRSSSLBind("ReportServerWebService")  
      Call SetRSSSLBind("ReportServerWebApp") 
    Case strSQLVersion >= "SQL2016"
      Call SetRSSSLBind("ReportServerWebService")  
      Call SetRSSSLBind("ReportServerWebApp") 
    Case Else
      Call SetRSSSLBind("ReportServerWebService")   
      Call SetRSSSLBind("ReportManager") 
  End Select

  strRSNamespace    = strStoreNamespace

End Sub


Private Sub SetRSSSLBind(strApplication)
  Call DebugLog("SetRSSSLBind: " & strApplication) 
' See https://community.certifytheweb.com/t/sql-server-reporting-services-ssrs/332.  Also IIS needs to be configured
  Dim strFunction

  strFunction        = SetRSInParam("CreateSSLCertificateBinding")
  objRSInParam.Properties_.Item("Application")                  = strApplication
  objRSInParam.Properties_.Item("CertificateHash")              = strSSLCertThumb
  objRSInParam.Properties_.Item("IPAddress")                    = "0.0.0.0"
  objRSInParam.Properties_.Item("Port")                         = strTCPPortSSL
  objRSInParam.Properties_.Item("Lcid")                         = intRSLcid
  Call RunRSWMI(strFunction, "-2147220932") ' OK if Binding already exists

End Sub


Private Sub SetRSWMI()
  Call DebugLog("SetRSWMI: " & strRSWMIPath)

  Set objRSWMI      = GetObject(strRSWMIPath)

  Select Case True
    Case strSQLVersion >= "SQL2017"
      intRSLcid     = 1033
    Case strSetupPowerBI = "YES"
      intRSLcid     = 1033
    Case Else
      strPath       = GetBuildfileValue("HKLMSQL") & GetBuildfileValue("InstRegRS") & "\Setup\Language"
      intRSLcid     = objShell.RegRead(strPath)
  End Select

End Sub


End Class


Function RunRSWMI(strFunction, strOK)
  RunRSWMI        = FBManageRSWMI.RunRSWMI(strFunction, strOK)
End Function

Sub SetRSDatabase(strServer, strRSDBName)
  Call FBManageRSWMI.SetRSDatabase(strServer, strRSDBName)
End Sub

Sub SetRSDirectory(strApplication, strDirectory)
  Call FBManageRSWMI.SetRSDirectory(strApplication, strDirectory)
End Sub

Function SetRSInParam(strFunction)
  SetRSInParam      = FBManageRSWMI.SetRSInParam(strFunction)
End Function

Sub SetRSSSL()
  Call FBManageRSWMI.SetRSSSL()
End Sub