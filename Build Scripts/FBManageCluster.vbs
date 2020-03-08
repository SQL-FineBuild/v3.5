'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageCluster.vbs  
'  Copyright FineBuild Team © 2018 - 2020.  Distributed under Ms-Pl License
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
Dim objCluster
Dim FBManageCluster: Set FBManageCluster = New FBManageClusterClass

Class FBManageClusterClass
  Dim objShell, objRE, objWMI, objWMIDNS, objWMIReg
  Dim strClusIPV4Address, strClusIPV4Mask, strClusIPV4Network, strClusIPV6Address, strClusIPV6Mask, strClusIPV6Network, strClusStorage, strClusterHost, strClusterName, strCmd, strCSVRoot
  Dim strFailoverClusterDisks, strOSVersion, strPath, strPathNew, strPreferredOwner, strServer, strSQLVersion, strUserDNSDomain, strUserDNSServer, strWaitLong, strWaitShort
  Dim intIndex


Private Sub Class_Initialize
  Call DebugLog("FBManageCluster Class_Initialize:")

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  Set objRE         = New RegExp
  objRE.Global      = True
  objRE.IgnoreCase  = True

  strClusIPV4Address  = GetBuildfileValue("ClusIPV4Address")
  strClusIPV4Mask     = GetBuildfileValue("ClusIPV4Mask")
  strClusIPV4Network  = GetBuildfileValue("ClusIPV4Network")
  strClusIPV6Address  = GetBuildfileValue("ClusIPV6Address")
  strClusIPV6Mask     = GetBuildfileValue("ClusIPV6Mask")
  strClusIPV6Network  = GetBuildfileValue("ClusIPV6Network")
  strClusStorage    = GetBuildfileValue("ClusStorage")
  strClusterHost    = ""
  strCSVRoot        = GetBuildfileValue("CSVRoot")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPreferredOwner = GetBuildfileValue("PreferredOwner")
  strServer         = GetBuildfileValue("AuditServer")
  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")

  strUserDNSDomain  = ""
  Set objWMIDNS     = Nothing

  objWMIReg.GetStringValue strHKLM,"Cluster\","ClusterName",strClusterName
  If strClusterName > "" Then
    strClusterHost  = "YES"
    Call OpenCluster()
  End If
  Call SetBuildfileValue("ClusterHost", strClusterHost)

End Sub


Sub BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strPriority)
  Call DebugLog("BuildCluster: " &strProcess)
  Dim strVolLabel, strVolName

  Call SetBuildfileValue("FailoverClusterDisks", "")

  Call SetupClusterGroup(strClusterGroup, strPriority)

  If strResourceName <> "" Then
    Call SetupClusterService(strClusterGroup, strResourceName, strServiceName, strServiceType, strServiceDesc, strServiceCheck)
  End If

  If strNetworkName <> "" Then
    Call SetupClusterNetwork(strProcess, strClusterGroup, strResourceName, strNetworkName)
  End If

  If strVolList <> "" Then
    Call MoveClusterVolume(strClusterGroup, strResourceName, strVolList)
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

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDOWNER:""" & strServer & """" 
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  Call SetResourceOn(strResourceName, "")

End Sub


Function GetClusterGroups()
  Call DebugLog("GetClusterGroups:")

  Set GetClusterGroups = objCluster.ResourceGroups

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


Function GetPrimaryNode(strGroup)
  Call DebugLog("GetPrimaryNode: " & strGroup)
  Dim colClusGroups, colClusResources, colOwnerNodes
  Dim objClusGroup, objClusResource, objOwnerNode
  Dim strPrimaryNode, strOwnerNode

  strPrimaryNode    = ""
  Set colClusGroups = GetClusterGroups()
  For Each objClusGroup In colClusGroups
    If objClusGroup.Name = strGroup Then
      strOwnerNode         = objClusGroup.OwnerNode.Name
      Set colClusResources = objClusGroup.Resources
      For Each objClusResource In colClusResources
        Set colOwnerNodes = objClusResource.PossibleOwnerNodes
        For Each objOwnerNode In colOwnerNodes
          Select Case True
            Case strOwnerNode = objOwnerNode.Name
              strPrimaryNode  = strOwnerNode
              Exit For
            Case Else
              strPrimaryNode  = objOwnerNode.Name
          End Select
        Next
        Exit For
      Next
      Exit For
    End If
  Next

  GetPrimaryNode    = strPrimaryNode

End Function


Sub MoveToGroup(strClusterGroup, strNode)
  Call DebugLog("MoveToGroup:" & strClusterGroup)
  Dim strNewNode

  Select Case True
    Case strNode = ""
      strNewNode    = strServer
    Case Else
      strNewNode    = strNode
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /MOVETO:""" & strNewNode & """ "
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

End Sub


Private Sub OpenCluster()
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
  Select Case True
    Case IsNull(objCluster)
      ' Nothing
    Case Else
      strClusterName = UCase(objCluster.Name)
  End Select
  Call SetBuildfileValue("ClusterName", strClusterName)
  Wscript.Sleep strWaitShort

End Sub


Sub RemoveOwner(strNetworkName)
  Call DebugLog("RemoveOwner: " & strNetworkName)

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Network Name (" & strNetworkName & ")"" /REMOVEOWNER:""" & strServer & """ "
  Call Util_RunExec(strCmd, "", strResponseYes, 5042)

End Sub



Sub SetOwnerNode(strClusterGroup)
  Call DebugLog("SetOwnerNode:" & strClusterGroup)

  If strPreferredOwner = strServer Then
    strCmd          = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /SETOWNERS:""" & strServer & """ "
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Call MoveToGroup(strClusterGroup, "")
  End If

End Sub


Private Sub SetupClusterGroup(strClusterGroup, strPriority)
  Call DebugLog("SetupClusterGroup: " & strClusterGroup)
  Dim strPriorityValue

  Select Case True
    Case strPriority = "L"
      strPriorityValue = 1000
    Case strPriority = "H"
      strPriorityValue = 3000
    Case Else
      strPriorityValue = 2000
  End Select

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /CREATE"
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  Call SetResourceOff(strClusterGroup, "GROUP")
  Call MoveToGroup(strClusterGroup, "")

  If strOSVersion >= "6.2" Then
    strCmd          = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroup & """ /PROP Priority=" & strPriorityValue
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

End Sub


Private Sub SetupClusterService(strClusterGroup, strResourceName, strServiceName, strServiceType, strServiceDesc, strServiceCheck)
  Call DebugLog("SetupClusterService:" & strResourceName & ", " & strServiceName)

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
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName & """ /ADDCHECK:""" & strServiceCheck & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 183)
  End If

End Sub


Private Sub SetupClusterNetwork(strProcess, strClusterGroup, strResourceName, strNetworkName)
  Call DebugLog("SetupClusterNetwork:")
  Dim strDNSName, strNetAddress

  strDNSName        = Left(strClusterGroup, Instr(strClusterGroup & " ", " ") - 1)
  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""Network Name"" /PRIV DNSNAME=""" & strDNSName & """"
  If strOSVersion < "6.3A" Then
    strCmd          = strCmd & " /PRIV NAME=""" & strDNSName & """"
  End If
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  If strClusIPV4Network <> "" Then
    strNetAddress   = GetClusterIPAddress(strNetworkName, strProcess, "IPv4", "IP")
    Call SetBuildfileValue("ClusterIPV4" & strProcess, strNetAddress)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE ""IP Address " & strDNSName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""IP Address""   /PRIV ADDRESS=" & strNetAddress & " SUBNETMASK=""" & strClusIPV4Mask & """ NETWORK=""" & strClusIPV4Network & """ ENABLENETBIOS=0" 
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /ADDDEP:""IP Address " & strDNSName & """"
    Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

  If strClusIPV6Network <> "" Then
    strNetAddress   = GetClusterIPAddress(strNetworkName, strProcess, "IPv6", "IP")
    Call SetBuildfileValue("ClusterIPV6" & strProcess, strNetAddress)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE ""IPv6 Address " & strDNSName & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""IPv6 Address"" /PRIV ADDRESS=" & strNetAddress & " NETWORK=""" & strClusIPV6Network & """ ENABLENETBIOS=0" 
    Call Util_RunExec(strCmd, "", strResponseYes, 5010)
  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strNetworkName & """ /ADDDEP:""IPv6 Address " & strDNSName & """"
  Call Util_RunExec(strCmd, "", strResponseYes, 5003)
  End If

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
      strRetAddress =  GetAddressPing(strAddress, strAddrType, strFormat)
   Case Else
      strRetAddress =  GetAddressWin32(strAddress, strAddrType, strFormat)
  End Select

  Select Case True
    Case strUserDNSServer = ""
      ' Nothing
    Case objWMIDNS Is Nothing
      strDebugMsg1  = "DNS Server: " & strUserDNSServer
      Set objWMIDNS = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strUserDNSServer & "\root\MicrosoftDNS")
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
    Case strRetAddress <> ""
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
  Dim objExec
  Dim strReadAll, strReadLine, strRetAddress

  strCmd            = "PING -a -n 1 " & strAddress
  Set objExec       = objShell.Exec(strCmd)
  strReadAll        = Replace(objExec.StdOut.ReadAll, vbLf, "")
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
    Case (strAddrType = "Server") And (intAddrPos > 0)
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
          Call SetBuildMessage(strMsgError, "IP Address exceeds maximum:" & intIP)
        End If
      Case Else
        intIP       = Hex(CLng("&h" & intIP) + 1)
        If intIP > "FFFF" Then
          Call SetBuildMessage(strMsgError, "IP Address exceeds maximum:" & intIP)
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


Private Sub MoveClusterVolume(strClusterGroup, strResourceName, strVolList)
  Call DebugLog("MoveClusterVolume: " & strVolList & " for " & strClusterGroup)
  Dim arrVolumes
  Dim strVolLabel, strVolParam, strVolSource, strVolType
  Dim intVol

  Call MoveToGroup(strClusStorage, "")

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
        strVolLabel = MoveClusterDrive(strClusterGroup, strVolParam)
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
        Call SetResourceOn(strVolRes, strResourceType)
    End Select
  Next

  MoveClusterCSV    = strVolRes

End Function


Private Function MoveClusterMP(strClusterGroup, strVolParam)
  Call DebugLog("MoveClusterMP: " & strVolParam & " for " & strClusterGroup)

  MoveClusterMP     = ""

End Function


Private Function MoveClusterDrive(strClusterGroup, strVolParam)
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
        strDebugMsg1            = "Moving " & strVolLabel & " to " & strClusterGroup
        strFailoverClusterDisks = strFailoverClusterDisks & """" & strVolLabel & """ "
        strCmd      = "CLUSTER """ & strClusterName & """ RESOURCE """ & strVolLabel & """ /MOVE:""" & strClusterGroup & """"
        Call Util_RunExec(strCmd, "", strResponseYes, 183)
        Call SetResourceOn(strVolLabel, "")
    End Select
  Next

  MoveClusterDrive  = strVolLabel

End Function


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
  If intErrSave <> 0 Then
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Wscript.Sleep strWaitLong
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

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


Sub BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strPriority)
  Call FBManageCluster.BuildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceType, strServiceName, strServiceDesc, strServiceCheck, strVolList, strPriority)
End Sub

Sub AddChildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceName, strServiceDesc, strServiceCheck)
  Call FBManageCluster.AddChildCluster(strProcess, strClusterGroup, strResourceName, strNetworkName, strServiceName, strServiceDesc, strServiceCheck)
End Sub

Sub AddChildNode(strProcess, strResourceName)
  Call FBManageCluster.AddChildNode(strProcess, strResourceName)
End Sub

Function GetAddress(strAddress, strFormat, strPreserve)
  GetAddress = FBManageCluster.GetAddress(strAddress, strFormat, strPreserve)
End Function

Function GetClusterGroups()
  Set GetClusterGroups  = FBManageCluster.GetClusterGroups()
End Function

Function GetClusterIPAddresses(strClusterGroup, strClusterType, strAddressFormat)
  GetClusterIPAddresses = FBManageCluster.GetClusterIPAddresses(strClusterGroup, strClusterType, strAddressFormat)
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

Sub MoveToGroup(strClusterGroup, strNode)
  Call FBManageCluster.MoveToGroup(strClusterGroup, strNode)
End Sub

Sub RemoveOwner(strNetworkName)
  Call FBManageCluster.RemoveOwner(strNetworkName)
End Sub

Sub SetOwnerNode(strCluster)
  Call FBManageCluster.SetOwnerNode(strCluster)
End Sub

Sub SetResourceOff(strResource, strResourceType)
  Call FBManageCluster.SetResourceOff(strResource, strResourceType)
End Sub

Sub SetResourceOn(strResource, strResourceType)
  Call FBManageCluster.SetResourceOn(strResource, strResourceType)
End Sub

Sub SetVolumeDependency(strResourceName, strVolParam)
  Call FBManageCluster.SetVolumeDependency(strResourceName, strVolParam)
End Sub