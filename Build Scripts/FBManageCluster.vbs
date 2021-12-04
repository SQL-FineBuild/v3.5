'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageCluster.vbs  
'  Copyright FineBuild Team © 2018 - 2021.  Distributed under Ms-Pl License
'
'  Purpose:      Cluster Management Utilities 
'
'  Author:       Ed Vassie
'
'  Date:         29 Jan 2018
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     29 Jan 2018  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageCluster: Set FBManageCluster = New FBManageClusterClass

Class FBManageClusterClass
  Dim colNodes, colResources
  Dim objCluster, objNode, objShell, objRE, objResource, objWMI, objWMIClus, objWMIDNS, objWMIReg
  Dim strClusIPV4Address, strClusIPV4Mask, strClusIPV4Network, strClusIPV6Address, strClusIPV6Mask, strClusIPV6Network
  Dim strClusterNetworkAO, strClusStorage, strClusterHost, strClusterName, strCmd, strCmdPS, strCSVRoot
  Dim strFailoverClusterDisks, strHKLM, strOSVersion, strPath, strPathNew, strPreferredOwner
  Dim strServer, strSQLVersion, strUserDNSDomain, strUserDNSServer, strWaitLong, strWaitShort
  Dim intIndex


Private Sub Class_Initialize
  Call DebugLog("FBManageCluster Class_Initialize:")

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  Set objRE         = New RegExp
  objRE.Global      = True
  objRE.IgnoreCase  = True

  strHKLM           = &H80000002
  strClusIPV4Address  = GetBuildfileValue("ClusIPV4Address")
  strClusIPV4Mask     = GetBuildfileValue("ClusIPV4Mask")
  strClusIPV4Network  = GetBuildfileValue("ClusIPV4Network")
  strClusIPV6Address  = GetBuildfileValue("ClusIPV6Address")
  strClusIPV6Mask     = GetBuildfileValue("ClusIPV6Mask")
  strClusIPV6Network  = GetBuildfileValue("ClusIPV6Network")
  strClusterNetworkAO = GetBuildfileValue("ClusterNetworkAO")
  strClusStorage    = GetBuildfileValue("ClusStorage")
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterName    = GetBuildfileValue("ClusterName")
  strCmdPS          = GetBuildfileValue("CmdPS")
  strCSVRoot        = GetBuildfileValue("CSVRoot")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPreferredOwner = GetBuildfileValue("PreferredOwner")
  strServer         = GetBuildfileValue("AuditServer")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strCSVRoot        = GetBuildfileValue("CSVRoot")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")

  strUserDNSDomain  = ""
  Set objWMIDNS     = Nothing

  If strClusterHost = "YES" Then
    Call ConnectCluster()
  End If

End Sub


Sub BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strVolReset, strPriority)
  Call DebugLog("BuildCluster: " &strProcess)
  Dim strVolLabel, strVolName

  Call SetBuildfileValue("FailoverClusterDisks", "")

  Call SetupClusterGroup(strClusterGroup)

  Call SetGroupPriority(strClusterGroup, strPriority)

  If strResourceName <> "" Then
    Call SetupClusterService(strClusterGroup, strResourceName, strServiceName, strServiceType, strServiceDesc, strServiceCheck)
  End If

  If strNetworkName <> "" Then
    Call SetupClusterNetwork(strProcess, strClusterGroup, strResourceName, strNetworkName)
  End If

  If strVolList <> "" Then
    Call MoveClusterVolume(strClusterGroup, strResourceName, strVolList, strVolReset)
  End If

End Sub


Sub AddChildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceName, strServiceDesc, strServiceCheck)
  Call DebugLog("AddChildCluster: " & strProcess)

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""Generic Service"" /PROP DESCRIPTION=""" & strServiceDesc & """"
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)
  Call SetResourceOff(strResourceName, "")

  If strResourceName <> "" Then
    Call SetupClusterService(strClusterGroup, strResourceName, strServiceName, "", strServiceDesc, strServiceCheck)
  End If

  If strNetworkName <> "" Then
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDDEP:""" & strNetworkName & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

End Sub


Sub AddChildNode(strProcess, strResourceName)
  Call DebugLog("AddChildNode: " & strProcess)

  Call SetResourceOff(strResourceName, "") 

  Call AddOwner(strResourceName)

  Call SetResourceOn(strResourceName, "")

End Sub


Sub AddOwner(strResourceName)
  Call DebugLog("AddOwner: " & strResourceName)

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDOWNER:""" & strServer & """" 
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

End Sub


Function CheckClusterHost()
  Call DebugLog("CheckClusterHost:")

  objWMIReg.GetStringValue strHKLM,"Cluster\","ClusterName",strClusterName
  Select Case True
    Case strClusterName > ""
      CheckClusterHost  = "YES"
    Case Else
      CheckClusterHost = ""
  End Select

End Function


Sub ConnectCluster()
  Call DebugLog("OpenCluster:")
  On Error Resume Next

  strClusterName    = ""
  Set objCluster    = CreateObject("MSCluster.Cluster")
  objCluster.Open ""
  Select Case True
    Case err.Number = 0
      Wscript.Sleep strWaitLong
    Case Else ' Network stack must be ready when Cluster Service starts, otherwise RPC error (often 1722) given.  Restart Cluster and wait so it can become ready.
      Call Util_RunExec("NET STOP  ""Cluster Service""", "", strResponseYes, 0)
      Wscript.Sleep strWaitLong
      Call Util_RunExec("NET START ""Cluster Service""", "", strResponseYes, 0)
      Wscript.Sleep strWaitLong
      Wscript.Sleep strWaitLong
      Wscript.Sleep strWaitLong
      objCluster.Open ""
  End Select
  intErrSave        = err.Number

  Wscript.Sleep strWaitShort
  Select Case True
    Case IsNull(objCluster)
      ' Nothing
    Case Else
      strClusterHost = "YES"
      strClusterName = UCase(objCluster.Name)
  End Select

  strOSVersion      = GetBuildfileValue("OSVersion")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  Set objWMIClus    = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\mscluster")
  Call SetBuildfileValue("ClusterHost", strClusterHost)
  Call SetBuildfileValue("ClusterName", strClusterName)

End Sub


Function GetClusterGroups()
  Call DebugLog("GetClusterGroups:")

  Set GetClusterGroups = objCluster.ResourceGroups

End Function


Function GetClusterInterfaces()
  Call DebugLog("GetClusterInterfaces:")

  Set GetClusterInterfaces = objCluster.NetInterfaces

End Function


Function GetClusterNetworks()
  Call DebugLog("GetClusterNetworks:")

  Set GetClusterNetworks = objCluster.Networks

End Function


Function GetClusterNodes()
  Call DebugLog("GetClusterNodes:")

  Set GetClusterNodes  = objCluster.Nodes

End Function


Function GetClusterResources()
  Call DebugLog("GetClusterResources:")

  Set GetClusterResources = objCluster.Resources

End Function


Function GetPrimaryNode(strResource)
  Call DebugLog("GetPrimaryNode: " & strResource)
  Dim colClusResources,colOwnerNodes
  Dim objClusResource,objOwnerNode
  Dim strPrimaryNode

  strPrimaryNode    = ""
  Set colClusResources = GetClusterResources()
  For Each objClusResource In colClusResources
    If UCase(objClusResource.Name) = UCase(strResource) Then
      Set colOwnerNodes = objClusResource.PossibleOwnerNodes
      For Each objOwnerNode In colOwnerNodes
        Call DebugLog("Resource: " & objClusResource.Name & ", Owner: " & objOwnerNode.Name)
        strPrimaryNode = UCase(objOwnerNode.Name)
        Exit For
      Next
    End If
    If strPrimaryNode <> "" Then
      Exit For
    End If
  Next

  GetPrimaryNode    = strPrimaryNode

End Function


Function GetResService(strResourceName)
  Dim strService

  Set colResources  = GetClusterResources()
  strService        = ""

  For Each objResource In colResources
    If UCase(objResource.Name) = UCase(strResourceName) Then
      strService    = objResource.PrivateProperties("ServiceName")
      Exit For
    End If
  Next

  GetResService = strService

End Function


Function GetStorageGroup(strGroupDefault)
  Call DebugLog("GetStorageGroup:")

  Select Case True
    Case strOSVersion < "6.2"
      GetStorageGroup = strGroupDefault
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & GetStorageGroup & """ /CREATE"
      Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    Case Else
      objWMIReg.GetStringValue strHKLM,"Cluster\","AvailableStorage",strPath
      objWMIReg.GetStringValue strHKLM,"Cluster\Groups\" & strPath & "\","Name",GetStorageGroup
  End Select

End Function


Sub MoveToNode(strClusterGroup, strNode)
  Call DebugLog("MoveToNode: " & strClusterGroup)
  Dim strNewNode

  Select Case True
    Case strNode = ""
      strNewNode    = strServer
    Case Instr("\", strNode) > 0
      strNewNode    = Left(strNode, Instr("\", strNode) - 1)
    Case Else
      strNewNode    = strNode
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /MOVETO:""" & strNewNode & """ "
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  If intErrSave <> 0 Then
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

End Sub


Sub RemoveOwner(strResourceName, strOwnerNode)
  Call DebugLog("RemoveOwner: " & strResourceName)
  Dim strOwner

  Select Case True
    Case strOwnerNode <> ""
      strOwner      = strOwnerNode
    Case Else
      strOwner      = strServer
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /REMOVEOWNER:""" & strOwner & """ "
  Call Util_RunExec(strCmd, "", strResponseYes, 5042)

End Sub


Sub SetGroupPriority(strClusterGroup, strPriority)
  Call DebugLog("SetGroupPriority: " & strPriority)
  Dim strPriorityValue

  Select Case True
    Case strPriority = "L"
      strPriorityValue = 1000
    Case strPriority = "H"
      strPriorityValue = 3000
    Case Else
      strPriorityValue = 2000
  End Select

  If strOSVersion >= "6.2" Then
    strCmd          = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /PROP Priority=" & strPriorityValue
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

End Sub


Sub SetOwnerNode(strClusterGroup)
  Call DebugLog("SetOwnerNode: " & strClusterGroup)

  If strPreferredOwner = strServer Then
    strCmd          = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /SETOWNERS:""" & strServer & """ "
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Call MoveToNode(strClusterGroup, "")
  End If

End Sub


Private Sub SetupClusterGroup(strClusterGroup)
  Call DebugLog("SetupClusterGroup: " & strClusterGroup)

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /CREATE"
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  Call SetResourceOff(strClusterGroup, "GROUP")
  Call MoveToNode(strClusterGroup, "")

End Sub


Private Sub SetupClusterService(strClusterGroup, strResourceName, strServiceName, strServiceType, strServiceDesc, strServiceCheck)
  Call DebugLog("SetupClusterService: " & strResourceName & ", " & strServiceName)
  Dim strResGUID, strServiceKey

  If strServiceType <> "" Then
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /CREATE /GROUP:""" & strClusterGroup & """ "
    strCmd          = strCmd & "/TYPE:""" & strServiceType & """ "
    If strServiceDesc <> "" Then
      strCmd        = strCmd & "/PROP DESCRIPTION=""" & strServiceDesc & """"
    End If
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
  End If

  If strServiceName <> "" Then
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /PROP RestartAction=""2"""
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /PRIV ServiceName=""" & strServiceName & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
  End If

  If strServiceCheck <> "" Then
    strServiceKey   = strServiceCheck
    If Instr(strServiceKey, "{GUID}") > 0 Then
      Set colResources = objWMIClus.ExecQuery("Select Id from MSCluster_Resource WHERE Name = '" & strResourceName & "'")
      For Each objResource In colResources
        strResGUID  = objResource.Id
      Next
      strServiceKey  = Replace(strServiceKey, "{GUID}", strResGUID)
      Call Util_RegWrite("HKLM\" & strServiceKey, "", "REG_SZ")
    End If
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDCHECK:""" & strServiceKey & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 183)
  End If

  strDebugMsg1      = "Remove Secondary nodes from Service Name ownership list"
  Set colNodes      = GetClusterNodes()
  For Each objNode In colNodes
    If UCase(objNode.Name) <> strServer Then
      Call RemoveOwner(strResourceName, objNode.Name)
    End If
  Next

End Sub


Private Sub SetupClusterNetwork(strProcess, strClusterGroup, strResourceName, strNetworkName)
  Call DebugLog("SetupClusterNetwork:")
  Dim strDNSName, strNetAddress

' Create Network Name
  strDNSName        = Left(strClusterGroup, Instr(strClusterGroup & " ", " ") - 1)
  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""Network Name"" /PRIV DNSNAME=""" & strDNSName & """"
  If strOSVersion < "6.3A" Then
    strCmd          = strCmd & " /PRIV NAME=""" & strDNSName & """"
  End If
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

' Add IPV4 Address
  If strClusIPV4Network <> "" Then
    strNetAddress   = GetClusterIPAddress(strNetworkName, strProcess, "IPv4", "IP")
    Call SetBuildfileValue("ClusterIPV4" & strProcess, strNetAddress)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE ""IP Address " & strDNSName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""IP Address""   /PRIV ADDRESS=" & strNetAddress & " SUBNETMASK=""" & strClusIPV4Mask & """ NETWORK=""" & strClusIPV4Network & """ ENABLENETBIOS=0" 
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /ADDDEP:""IP Address " & strDNSName & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

' Add IPV6 Address
  If strClusIPV6Network <> "" Then
    strNetAddress   = GetClusterIPAddress(strNetworkName, strProcess, "IPv6", "IP")
    Call SetBuildfileValue("ClusterIPV6" & strProcess, strNetAddress)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE ""IPv6 Address " & strDNSName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""IPv6 Address"" /PRIV ADDRESS=" & strNetAddress & " NETWORK=""" & strClusIPV6Network & """ ENABLENETBIOS=0" 
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /ADDDEP:""IPv6 Address " & strDNSName & """"
  Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

' Add Network Name Dependency
  If strResourceName <> "" Then
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDDEP:""" & strNetworkName & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

End Sub


Function GetClusterIPAddresses(strClusterGroup, strClusterType, strAddressFormat)
  Call DebugLog("GetClusterIPAddresses: " & strClusterGroup)
  Dim strFailoverClusterIPAddresses, strClusterIPExtra, strClusterIPV4, strClusterIPV6

  strFailoverClusterIPAddresses = ""
  strClusterIPExtra             = GetBuildfileValue("Clus" & strClusterType & "IPExtra")
  strClusterIPV4                = ""
  strClusterIPV6                = ""

  If strClusIPV4Network <> "" Then
    strClusterIPV4  = GetClusterIPAddress(strClusterGroup, strClusterType, "IPv4", strAddressFormat)
    strFailoverClusterIPAddresses = strFailoverClusterIPAddresses & strClusterIPV4
  End If

  If strClusIPV6Network <> "" Then
    strClusterIPV6  = GetClusterIPAddress(strClusterGroup, strClusterType, "IPv6", strAddressFormat)
    strFailoverClusterIPAddresses = strFailoverClusterIPAddresses & strClusterIPV6
  End If

  If strClusterIPExtra <> "" Then
    strFailoverClusterIPAddresses = strFailoverClusterIPAddresses & " " & strClusterIPExtra
  End If

  Call SetBuildfileValue("ClusterIPV4" & strClusterType, strClusterIPV4)
  Call SetBuildfileValue("ClusterIPV6" & strClusterType, strClusterIPV6)
  GetClusterIPAddresses = strFailoverClusterIPAddresses

End Function


Private Function GetClusterIPAddress(strClusterName, strClusType, strIPType, strAddrType)
  Call DebugLog("GetClusterIPAddress: " & strClusterName)
  Dim strClusIPAddr, strReturnAddress

  strClusIPAddr     = GetAddress(strClusterName, "IP", "")
  If strClusIPAddr = "" Then
    strClusIPAddr   = GetNextAddress(strClusterName, strClusType, strIPType, strAddrType)
  End If

  Select Case True
    Case strAddrType = "IP"
      strReturnAddress = strClusIPAddr
    Case strSQLVersion = "SQL2005" 
      strReturnAddress = strClusIPAddr & "," & strClusIPV4Network
    Case (strAddrType = "SET") And (strIPType = "IPv4")
      strReturnAddress = "('" & strClusIPAddr & "','" & strClusIPV4Mask & "')"
    Case strAddrType = "SET"
      strReturnAddress = "('" & strClusIPAddr & "')"
    Case strIPType = "IPv4"
      strReturnAddress = strIPType & ";" & strClusIPAddr & ";" & strClusIPV4Network & ";" & strClusIPV4Mask
    Case Else
      strReturnAddress = strIPType & ";" & strClusIPAddr & ";" & strClusIPV6Network 
  End Select

  If strAddrType <> "SET" Then
    strReturnAddress = """" & strReturnAddress & """ "
  End If

  GetClusterIPAddress  = strReturnAddress

End Function


Function GetAddress(strAddress, strFormat, strPreserve)
  Call DebugLog("GetAddress: " & strAddress)
  Dim strAddrType, strQuery, strRetAddress

  If strUserDNSDomain = "" Then
    strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
    strUserDNSServer  = GetBuildfileValue("UserDNSServer")
    strOSVersion      = GetBuildfileValue("OSVersion")
    strWaitLong       = GetBuildfileValue("WaitLong")
    If strUserDNSServer <> "" Then
      strDebugMsg1    = "DNS Server: " & strUserDNSServer
      Set objWMIDNS   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strUserDNSServer & "\root\MicrosoftDNS")
    End If
  End If

  objRE.Pattern     = "^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$"
  Select Case True
    Case objRE.Test(strAddress)
      strAddrType   = "IPv4"
    Case Instr(strAddress, ":") > 0
      strAddrType   = "IPv6"
    Case Else
      strAddrType   = "Server"
  End Select

  Select Case True
    Case strOSVersion < "6.0"
      strRetAddress =  GetAddressPing(strAddress,  strAddrType, strFormat)
   Case Else
      strRetAddress =  GetAddressWin32(strAddress, strAddrType, strFormat)
  End Select

  Select Case True
    Case strRetAddress <> ""
      ' Nothing
    Case objWMIDNS Is Nothing
      If strAddrType = "Server" Then
        strRetAddress = strServer
      End If
    Case strAddrType = "IPv4"
      strQuery      = "SELECT * FROM MicrosoftDNS_AType     WHERE IPAddress   = """ & strAddress & """"
      strRetAddress = GetAddressDNS(strQuery, strAddrType, strFormat)
    Case strAddrType = "IPv6"
      strQuery      = "SELECT * FROM MicrosoftDNS_AAAAType  WHERE IPV6Address = """ & strAddress & """"
      strRetAddress = GetAddressDNS(strQuery, strAddrType, strFormat)
    Case UCase(strFormat) = "ALIAS"
      strQuery      = "SELECT * FROM MicrosoftDNS_CNAMEType WHERE OwnerName   = """ & strAddress & "." & strUserDNSDomain & """"
      strRetAddress = GetAddressDNS(strQuery, strAddrType, strFormat)
    Case Else
      strQuery      = "SELECT * FROM MicrosoftDNS_AType     WHERE OwnerName   = """ & strAddress & "." & strUserDNSDomain & """"
      strRetAddress = GetAddressDNS(strQuery, strAddrType, strFormat)
      If strRetAddress = "" Then
        strQuery    = "SELECT * FROM MicrosoftDNS_AAAAType  WHERE OwnerName   = """ & strAddress & "." & strUserDNSDomain & """"
        strRetAddress = GetAddressDNS(strQuery, strAddrType, strFormat)
      End If
  End Select

  Call DebugLog("Address found: """ & strRetAddress & """")
  Select Case True
    Case strRetAddress = ""
      ' Nothing
    Case strPreserve = "Y"
      strRetAddress = strAddress
  End Select 

  GetAddress     = UCase(strRetAddress)

End Function


Private Function GetAddressWin32(strAddress, strAddrType, strFormat)
  Call DebugLog("GetAddressWin32:")
  Dim objAddr
  Dim strRetAddress

  strRetAddress     = ""
  Set objAddr       = objWMI.Get("Win32_PingStatus.Address='" & strAddress & "',ResolveAddressNames=True,TypeOfService=4")
  Select Case True
    Case objAddr.StatusCode = 0
      strRetAddress =  objAddr.ProtocolAddress
    Case Else
      strRetAddress =  GetAddressPing(strAddress, strAddrType, strFormat)
  End Select

  GetAddressWin32   = strRetAddress

End Function


Private Function GetAddressPing(strAddress, strAddrType, strFormat)
  Call DebugLog("GetAddressPing:")
  Dim arrReadAll
  Dim colAddrs
  Dim intLines, intAddrPos
  Dim strReadAll, strReadLine, strRetAddress

  strReadAll        = ExecPing("PING -a -n 1 -4 " & strAddress)
  If Instr(strReadAll, "[") = 0 Then
    strReadAll      = ExecPing("PING -a -n 1 " & strAddress)
  End If
  arrReadAll        = Filter(Split(strReadAll, vbCr), " ")
  intLines          = UBound(arrReadAll)
  Call DebugLog("PING output:" & Cstr(intLines) & ">" & Join(arrReadAll, "< >") & "<")

  strRetAddress     = ""
  strReadLine       = arrReadAll(0)
  intAddrPos        = Instr(strReadLine, "[")
  Select Case True
    Case intLines = 0
      ' Nothing
    Case intAddrPos = 0
       ' Nothing
    Case strFormat = "IP"
      strRetAddress = Mid(strReadLine,  intAddrPos + 1)
      strRetAddress = Left(strRetAddress, Instr(strRetAddress, "]") - 1)
    Case UCase(strFormat) = "ALIAS"
      strRetAddress = Mid(strReadLine,  intAddrPos + 1)
      strRetAddress = Left(strRetAddress, Instr(strRetAddress, "]") - 1)
    Case strAddrType = "Server"
      strRetAddress = Mid(strReadLine,  intAddrPos + 1)
      strRetAddress = Left(strRetAddress, Instr(strRetAddress, "]") - 1)
    Case Else
      strRetAddress = Left(strReadLine, intAddrPos - 2)
      strRetAddress = Mid(strRetAddress, InstrRev(strRetAddress, " ") + 1)
      If Right(UCase(strRetAddress), Len(strUserDNSDomain)) = UCase(strUserDNSDomain) Then
        strRetAddress = Left(strRetAddress, Len(strRetAddress) - (Len(strUserDNSDomain) + 1))
      End If
  End Select

  GetAddressPing    = strRetAddress

End Function


Function ExecPing(strCmd)
  Call DebugLog("ExecPing: " & strCmd)
  Dim objExec

  Set objExec       = objShell.Exec(strCmd)
  ExecPing          = Replace(objExec.StdOut.ReadAll, vbLf, "")

End Function


Private Function GetAddressDNS(strQuery, strAddrType, strFormat)
  Call DebugLog("GetAddressDNS: " & strQuery)
  Dim colAddrs
  Dim objAddr
  Dim strRetAddress

  strRetAddress     = ""
  strDebugMsg1      = "Query: " & strQuery
  Set colAddrs      = objWMIDNS.ExecQuery(strQuery)
  If colAddrs.Count > 0 Then
    For Each objAddr In colAddrs
      strDebugMsg2  = "Addr: " & objAddr.OwnerName
      Select Case True
        Case (strFormat = "IP") And (strAddrType = "IPv6")
          strRetAddress = objAddr.IPV6Address
        Case strFormat = "IP"
          strRetAddress = objAddr.IPAddress
        Case UCase(strFormat) = "ALIAS"
          strRetAddress = objAddr.PrimaryName
          If Instr(strRetAddress, ".") > 0 Then
            strRetAddress = Left(strRetAddress, Instr(strRetAddress, ".") - 1)
          End If
        Case Else
          strRetAddress = objAddr.OwnerName
          If Instr(strRetAddress, ".") > 0 Then
            strRetAddress = Left(strRetAddress, Instr(strRetAddress, ".") - 1)
          End If
      End Select        
    Next
  End If

  GetAddressDNS     = strRetAddress

End Function


Private Function GetNextAddress(strClusterName, strClusType, strIPType, strAddrType)
  Call DebugLog("GetNextAddress: " & strClusterName)
  Dim intIP
  Dim strBaseIPAddr, strClusIPAddr, strClusIPIdx, strClusIPSuf

  Select Case True
    Case strIPType = "IPv4"
      intIndex      = InstrRev(strClusIPV4Address, ".")
      intIP         = Mid(strClusIPV4Address, intIndex + 1)
      strBaseIPAddr = Left(strClusIPV4Address, intIndex)
    Case Else
      intIndex      = InstrRev(strClusIPV6Address, ":")
      intIP         = Mid(strClusIPV6Address, intIndex + 1)
      If InstrRev(intIP, ".") > 0 Then
        intIndex    = InstrRev(strClusIPV6Address, ".")
        intIP       = Mid(strClusIPV6Address, intIndex + 1)
      End If
      strBaseIPAddr = Left(strClusIPV6Address, intIndex)
  End Select
  strClusIPSuf      = GetBuildfileValue("Clus" & strClusType & "IPSuffix")
  strClusIPAddr     = MergeAddrSuffix(strIPType, strBaseIPAddr, strClusIPSuf)

  While strClusIPSuf = ""
    Select Case True
      Case strIPType = "IPv4"
        intIP       = CStr(CInt(intIP) + 1)
        If intIP > 255 Then
          Call SetBuildMessage(strMsgError, "IP Address exceeds maximum: " & intIP)
        End If
      Case Else
        intIP       = Hex(CLng("&h" & intIP) + 1)
        If intIP > "FFFF" Then
          Call SetBuildMessage(strMsgError, "IP Address exceeds maximum: " & intIP)
        End If
    End Select
    strClusIPAddr   = strBaseIPAddr & intIP
    Select Case True
      Case GetAddress(strClusIPAddr, "IP", "") <> ""
        ' Nothing
      Case CheckAddressUsed(strClusterName, strClusIPAddr) = False
        strClusIPSuf = intIP
    End Select
  WEnd

  GetNextAddress    = strClusIPAddr

End Function


Private Function MergeAddrSuffix(strIPType, strBaseAddr, strSuffix)
  Call DebugLog("MergeAddrSuffix:")
  Dim arrBase, arrSuffix
  Dim intBase, intSuffix
  Dim strMergeAddr

  Select Case True
    Case strIPType = "IPv4"
      arrBase       = Split(strBaseAddr, ".")
      intBase       = UBound(arrBase)
      arrSuffix     = Split(strSuffix, ".")
      intSuffix     = UBound(arrSuffix)
      Select Case True
        Case intBase > 3
          strMergeAddr = strBaseAddr
        Case Else
          If intBase < 3 Then
            intBase = 3
            Redim Preserve arrBase(intBase)
          End If
          While intSuffix >= 0
            arrBase(intBase) = arrSuffix(intSuffix)
            intBase     = intBase - 1
            intSuffix   = intSuffix - 1
          WEnd
          strMergeAddr  = Join(arrBase, ".")
      End Select
    Case Else
      strMergeAddr  = strBaseAddr & strSuffix
  End Select

  MergeAddrSuffix   = strMergeAddr

End Function


Private Function CheckAddressUsed(strClusterName, strClusIPAddr)
  Call DebugLog("CheckAddressUsed: " & strClusIPAddr)
  Dim arrResources
  Dim objResource
  Dim strAddress, strResType

  CheckAddressUsed  = False
  strPath           = "Cluster\Resources"
  objWMIReg.EnumKey strHKLM, strPath, arrResources
  For Each objResource In arrResources
    strAddress      = ""
    strPathNew      = strPath & "\" & objResource
    objWMIReg.GetStringValue strHKLM, strPathNew, "Type", strResType
    Select Case True
      Case strResType = "IP Address"
        objWMIReg.GetStringValue strHKLM, strPathNew & "\Parameters", "Address", strAddress
      Case strResType = "IPV6 Address"
        objWMIReg.GetStringValue strHKLM, strPathNew & "\Parameters", "Address", strAddress
      Case strResType = "IPV6 Tunnel Address"
        objWMIReg.GetStringValue strHKLM, strPathNew & "\Parameters", "Address", strAddress
    End Select
    Select Case True
      Case IsNull(strAddress)
        ' Nothing
      Case strAddress = strClusIPAddr
        CheckAddressUsed = True
        Exit For
   End Select
  Next

End Function


Private Sub MoveClusterVolume(strClusterGroup, strResourceName, strVolList, strVolReset)
  Call DebugLog("MoveClusterVolume: " & strVolList & " for " & strClusterGroup)
  Dim arrVolumes
  Dim strVolLabel, strVolParam, strVolSource, strVolType
  Dim intVol

  Call MoveToNode(strClusStorage, "")

  arrVolumes        = Split(strVolList, " ")
  strFailoverClusterDisks = GetBuildfileValue("FailoverClusterDisks")
  For intVol = 0 To UBound(arrVolumes)
    strVolParam     = Trim(arrVolumes(intVol))
    strVolLabel     = ""
    strVolSource    = GetBuildfileValue(strVolParam & "Source")
    strVolType      = GetBuildfileValue(strVolParam & "Type")
    Select Case True
      Case strVolSource = "C"
        strVolLabel = MoveClusterCSV(strClusterGroup, strVolParam)
      Case (strVolSource = "D") And (strVolType <> "L")
        strVolLabel = MoveClusterDrive(strClusterGroup, strVolParam, strVolReset)
      Case strVolSource = "M"
        strVolLabel = MoveClusterMP(strClusterGroup, strVolParam)
    End Select
    If strResourceName <> "" Then
      Call SetVolumeDependency(strResourceName, strVolParam)
    End If
  Next

  Call SetBuildfileValue("FailoverClusterDisks", strFailoverClusterDisks)

End Sub


Private Function MoveClusterCSV(strClusterGroup, strVolParam)
  Call DebugLog("MoveClusterCSV: " & strVolParam & " for " & strClusterGroup)
  Dim arrItems
  Dim intIdx, intUBound
  Dim strVol, strVolRes, strVolList

  strVolList        = GetBuildfileValue(strVolParam)
  arrItems          = Split(strVolList, ",")
  intUBound         = UBound(arrItems)

  For intIdx = 0 To intUBound
    strVol          = arrItems(intIdx)
    strVol          = Mid(strVol, Len(strCSVRoot) + 1)
    If Instr(strVol, "\") > 0 Then
      strVol        = Left(strVol, Instr(strVol, "\") - 1)
    End If
    strVolRes       = GetBuildFileValue("Vol_" & UCase(strVol) & "Res")
    Select Case True
      Case Instr(strFailoverClusterDisks, """" & strVol & """") > 0
        ' Nothing
      Case Else
        strDebugMsg1            = "Moving " & strVolRes & " to " & strClusterGroup
        strFailoverClusterDisks = strFailoverClusterDisks & """" & strVolRes & """ "
        Call SetResourceOn(strVolRes, "")
    End Select
  Next

  MoveClusterCSV    = strVolRes

End Function


Private Function MoveClusterMP(strClusterGroup, strVolParam)
  Call DebugLog("MoveClusterMP: " & strVolParam & " for " & strClusterGroup)

  MoveClusterMP     = ""

End Function


Private Function MoveClusterDrive(strClusterGroup, strVolParam, strVolReset)
  Call DebugLog("MoveClusterDrive: " & strVolParam & " for " & strClusterGroup)
  Dim intIdx, intLen
  Dim strVol, strVolLabel, strVolList

  strVolList        = GetBuildfileValue(strVolParam)
  intLen            = Len(strVolList)
  For intIdx = 1 To intLen
    strVol          = Mid(strVolList, intIdx, 1)
    strVolLabel     = GetBuildFileValue("Vol" & strVol & "Label")
    Select Case True
      Case Instr(strFailoverClusterDisks, """" & strVolLabel & """") > 0
        ' Nothing
      Case Else
        Call SetResourceOff(strVolLabel, "")
        If strVolReset = "Y" Then
          Call ClearResParent(strVolLabel)
        End If
        strDebugMsg1            = "Moving " & strVolLabel & " to " & strClusterGroup
        strFailoverClusterDisks = strFailoverClusterDisks & """" & strVolLabel & """ "
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strVolLabel & """ /MOVE:""" & strClusterGroup & """"
        Call Util_RunExec(strCmd, "", strResponseYes, 183)
        Call SetResourceOn(strVolLabel, "")
    End Select
  Next

  MoveClusterDrive  = strVolLabel

End Function


Private Sub ClearResParent(strVolLabel)
  Call DebugLog("ClearResParent: " & strVolLabel)
  Dim colDependencies
  Dim objDependency

  Set colResources  = GetClusterResources()
  For Each objResource In colResources
    Set colDependencies = objResource.Dependencies
    For Each objDependency In colDependencies
      If objDependency.Name = strVolLabel Then
        strDebugMsg1 = "Clear " & strVolLabel & " From " & objResource.Name
        colDependencies.RemoveItem strVolLabel
        Call SetBuildfileValue("ResParent", objResource.Name)
      End If
   Next
  Next

End Sub


Sub SetClusNetworkProps(strClusterNetwork, strClusterAction)
  Call DebugLog("SetClusNetworkProps: " & strClusterNetwork)
  Dim colClusNodes
  Dim objClusNode

  Set colClusNodes  = GetClusterNodes()
  For Each objClusNode In colClusNodes
    Select Case True
      Case strClusterAction = "ADDNODE" 
        If UCase(strServer) = UCase(objClusNode.Name) Then
          Call SetResourceOff(strClusterNetwork, "")
          Call AddOwner(strClusterNetwork)
        End If
      Case UCase(strServer) = UCase(objClusNode.Name)
        Call SetResourceOff(strClusterNetwork, "")
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PROP PendingTimeout=600000:DWORD "
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV PublishPTRRecords=1 "
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV RequireDNS=1:DWORD "
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV RequireKerberos=1:DWORD "
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        If strClusterNetwork = strClusterNetworkAO Then
          strCmd    = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV RegisterAllProvidersIP = 0"
          Call Util_RunExec(strCmd, "", strResponseYes, 5024)
          strCmd    = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV HostRecordTTL = 300"
          Call Util_RunExec(strCmd, "", strResponseYes, 5024)
        End If
        If strOSVersion >= "6" Then
          strCmd    = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNetwork & """ /PRIV DeleteVcoOnResCleanup=1:DWORD "
          Call Util_RunExec(strCmd, "", strResponseYes, 0)
        End If
      Case Else
        Call SetResourceOff(strClusterNetwork, "")
        Call RemoveOwner(strClusterNetwork, objClusNode.Name)
    End Select
  Next

  Set colClusNodes  = Nothing

End Sub


Sub SetClusterCmd()
  Call DebugLog("SetClusterCmd:")
  Dim strStatusComplete

  strCmdPS          = GetBuildfileValue("CmdPS")
  strStatusComplete = GetBuildfileValue("StatusComplete")
  Select Case True
    Case GetBuildfileValue("SetupClusterCmdStatus") = strStatusComplete
      ' Nothing
    Case strOSVersion < "6.2" 
      ' Nothing
    Case Else
      strCmd        = strCmdPS & " -Command Install-WindowsFeature -Name RSAT-Clustering-AutomationServer"
      Call Util_RunExec(strCmd, "", "", 0)
      strCmd        = strCmdPS & " -Command Install-WindowsFeature -Name RSAT-Clustering-CmdInterface"
      Call Util_RunExec(strCmd, "", "", 0)
  End Select

  Call SetBuildfileValue("SetupClusterCmdStatus", strStatusComplete)

End Sub


Sub SetResourceOff(strResource, strResourceType)
  Call DebugLog("SetResourceOff: " & strResource)
  Dim strType
  
  Select Case True
    Case strResourceType = ""
      strType       = "RESOURCE"
    Case Else
      strType       = strResourceType
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ " & strType & " """ & strResource & """ /OFF"
  Call Util_RunExec(strCmd, "", strResponseYes, "5064")

End Sub


Sub SetResourceOn(strResource, strResourceType)
  Call DebugLog("SetResourceOn: " & strResource)
  Dim strType
  
  Select Case True
    Case strResourceType = ""
      strType       = "RESOURCE"
    Case Else 
      strType       = strResourceType
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ " & strType & " """ & strResource & """ /ON"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
'    Case (intErrSave = 5023) Or (intErrSave = 5942) ' Only needed if Cluster is already broken
      ' Nothing
    Case intErrSave = 5063
      ' Nothing
    Case Else
      Call DebugLog("Retrying due to code " & CStr(intErrSave))
      Wscript.Sleep strWaitLong
      Wscript.Sleep strWaitLong
      Wscript.Sleep strWaitLong
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

End Sub


Sub SetResourceAllOn()
  Call DebugLog("SetResourceAllOn:")

  Set colResources  = GetClusterResources()
  For Each objResource In colResources
    Select Case True
      Case objResource.State = 0 ' Resource Inherited
        ' Nothing
      Case objResource.State = 2 ' Resource Operational
        ' Nothing
      Case Else
        Call SetResourceOn(objResource.Name, "")
    End Select
  Next

End Sub


Sub SetVolumeDependency(strResourceName, strVolParam)
  Call DebugLog("SetVolumeDependency: " & strVolParam & " for " & strResourceName)
  Dim strVolName, strVolLabel

  If GetBuildfileValue(strVolParam & "Source") <> "C" Then
    strVolName    = GetBuildFileValue(strVolParam)
    strVolLabel   = GetBuildFileValue("Vol" & strVolName & "Label")
    strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDDEP:""" & strVolLabel & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

End Sub


End Class


Sub AddOwner(strResourceName)
  Call FBManageCluster.AddOwner(strResourceName)
End Sub

Sub ConnectCluster()
  Call FBManageCluster.ConnectCluster()
End Sub

Sub BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strVolReset, strPriority)
  Call FBManageCluster.BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strVolReset, strPriority)
End Sub

Sub AddChildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceName, strServiceDesc, strServiceCheck)
  Call FBManageCluster.AddChildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceName, strServiceDesc, strServiceCheck)
End Sub

Sub AddChildNode(strProcess, strResourceName)
  Call FBManageCluster.AddChildNode(strProcess, strResourceName)
End Sub

Function CheckClusterHost()
  CheckClusterHost  = FBManageCluster. CheckClusterHost()
End Function

Function GetAddress(strAddress, strFormat, strPreserve)
  GetAddress        = FBManageCluster.GetAddress(strAddress, strFormat, strPreserve)
End Function

Function GetClusterGroups()
  Set GetClusterGroups  = FBManageCluster.GetClusterGroups()
End Function

Function GetClusterIPAddresses(strClusterGroup, strClusterType, strAddressFormat)
  GetClusterIPAddresses = FBManageCluster.GetClusterIPAddresses(strClusterGroup, strClusterType, strAddressFormat)
End Function

Function GetClusterInterfaces()
  Set GetClusterInterfaces = FBManageCluster.GetClusterInterfaces()
End Function

Function GetClusterNetworks()
  Set GetClusterNetworks = FBManageCluster.GetClusterNetworks()
End Function

Function GetClusterNodes()
  Set GetClusterNodes    = FBManageCluster.GetClusterNodes()
End Function

Function GetClusterResources()
  Set GetClusterResources = FBManageCluster.GetClusterResources()
End Function

Function GetPrimaryNode(objResource)
  GetPrimaryNode    = FBManageCluster.GetPrimaryNode(objResource)
End Function

Function GetResService(strResourceName)
  GetResService     = FBManageCluster.GetResService(strResourceName)
End Function

Function GetStorageGroup(strGroupDefault)
  GetStorageGroup   = FBManageCluster.GetStorageGroup(strGroupDefault)
End Function

Sub MoveToNode(strClusterGroup, strNode)
  Call FBManageCluster.MoveToNode(strClusterGroup, strNode)
End Sub

Sub RemoveOwner(strResourceName, strOwnerNode)
  Call FBManageCluster.RemoveOwner(strResourceName, strOwnerNode)
End Sub

Sub SetGroupPriority(strClusterGroup, strPriority)
  Call FBManageCluster.SetGroupPriority(strClusterGroup, strPriority)
End Sub

Sub SetOwnerNode(strCluster)
  Call FBManageCluster.SetOwnerNode(strCluster)
End Sub

Sub SetClusNetworkProps(strClusterNetwork, strClusterAction)
  Call FBManageCluster.SetClusNetworkProps(strClusterNetwork, strClusterAction)
End Sub

Sub SetClusterCmd()
  Call FBManageCluster.SetClusterCmd()
End Sub

Sub SetResourceOff(strResource, strResourceType)
  Call FBManageCluster.SetResourceOff(strResource, strResourceType)
End Sub

Sub SetResourceOn(strResource, strResourceType)
  Call FBManageCluster.SetResourceOn(strResource, strResourceType)
End Sub

Sub SetResourceAllOn()
  Call FBManageCluster.SetResourceAllOn()
End Sub

Sub SetVolumeDependency(strResourceName, strVolParam)
  Call FBManageCluster.SetVolumeDependency(strResourceName, strVolParam)
End Sub