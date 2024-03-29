'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageSecurity.vbs  
'  Copyright FineBuild Team � 2020 - 2021.  Distributed under Ms-Pl License
'
'  Purpose:      Manage the FineBuild Account processing 
'
'  Author:       Ed Vassie
'
'  Date:         23 Apr 2020
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     23 Apr 2020  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBManageSecurity: Set FBManageSecurity = New FBManageSecurityClass

Class FBManageSecurityClass
Dim objADOCmd, objADOConn, objExec, objFolder, objFSO, objFW, objFWRules, objRecordSet, objSDUtil, objShell, objWMIReg
Dim arrProfFolders, arrProfUsers
Dim intIdx, intBuiltinDomLen, intNTAuthLen, intServerLen
Dim strBuiltinDom, strClusterName, strCmd, strCmdSQL, strDirSystemDataBackup, strGroupDBA, strGroupDBANonSA, strHKLM, strHKU, strIsInstallDBA, strLocalAdmin
Dim strNTAuth, strOSVersion, strPath, strProfDir, strProgCacls, strProgReg, strServer, strSIDDistComUsers, strUser, strUserAccount, strWaitShort


Private Sub Class_Initialize
  Call DebugLog("FBManageSecurity Class_Initialize:")

  Set objADOConn    = CreateObject("ADODB.Connection")
  Set objADOCmd     = CreateObject("ADODB.Command")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objFW         = CreateObject("HNetCfg.FwPolicy2")
  Set objFWRules    = objFW.Rules
  Set objSDUtil     = CreateObject("ADsSecurityUtility")
  Set objShell      = CreateObject("Wscript.Shell")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strHKLM           = &H80000002
  strHKU            = &H80000003
  strBuiltinDom     = GetBuildfileValue("BuiltinDom")
  strClusterName    = GetBuildfileValue("ClusterName")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strIsInstallDBA   = GetBuildfileValue("IsInstallDBA")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strNTAuth         = GetBuildfileValue("NTAuth")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strProgReg        = GetBuildfileValue("ProgReg")
  strServer         = GetBuildfileValue("AuditServer")
  strSIDDistComUsers  = GetBuildfileValue("SIDDistComUsers")
  strUserAccount    = GetBuildfileValue("UserAccount")
  strWaitShort      = GetBuildfileValue("WaitShort")

  Set arrProfFolders  = objFSO.GetFolder(strProfDir).SubFolders
  objWMIReg.EnumKey strHKU, "", arrProfUsers

  objADOConn.Provider            = "ADsDSOObject"
  objADOConn.Open "ADs Provider"
  Set objADOCmd.ActiveConnection = objADOConn

  intBuiltinDomLen  = Len(strBuiltinDom) + 1
  intNTAuthLen      = Len(strNTAuth) + 1
  intServerLen      = Len(strServer) + 1

End Sub


Sub BackupDBMasterKey(strDB, strPassword)
  Call DebugLog("BackupDBMasterKey: " & strDB)
  Dim strPathNew

  strPathNew        = strDirSystemDataBackup & "\" & strDB & "DBMasterKey.snk"
  Call DeleteFile(strPathNew)

  Call Util_ExecSQL(strCmdSQL & "-d """ & strDB & """ -Q", """BACKUP MASTER KEY TO FILE='" & strPathNew & "' ENCRYPTION BY PASSWORD='" & strPassword & "';""", 0)

End Sub


Private Function CheckGroup(strName, strUserDNSDomain)
  Call DebugLog("CheckGroup " & strName)
  Dim strAccountName

  Select Case True
    Case Instr(strName, "\") > 0
      strAccountName = Mid(strName, Instr(strName, "\") + 1)
    Case Instr(strName, "@") > 0
      strAccountName = Mid(1, strName, Instr(strName, "@") - 1)
    Case Else
      strAccountName = strName
  End Select

  objADOCmd.CommandText = "SELECT Name FROM 'LDAP://" & strUserDNSDomain & "' WHERE objectCategory='group' " & "AND Name='" & strAccountName & "'"
  Set objRecordSet = objADOCmd.Execute

  Select Case True
    Case objRecordSet.BOF
      CheckGroup    = False
    Case Else
      CheckGroup    = True
  End Select

End Function

Function FormatAccount(strAccount)
  Call DebugLog("FormatAccount: " & strAccount)
  Dim strFmtAccount

  Select Case True
    Case Left(strAccount, intNTAuthLen) = strNTAuth & "\"
      strFmtAccount = Mid(strAccount, intNTAuthLen + 1)
    Case Left(strAccount, intServerLen) = strServer & "\"
      strFmtAccount = Mid(strAccount, intServerLen + 1)
    Case Left(strAccount, intBuiltinDomLen) = strBuiltinDom & "\"
      strFmtAccount = Mid(strAccount, intBuiltinDomLen + 1)
    Case Else
      strFmtAccount = strAccount
  End Select

  Select Case True
    Case strFmtAccount = strServer
      strFmtAccount = strFmtAccount & "$"
    Case strFmtAccount = strClusterName
      strFmtAccount = strFmtAccount & "$"
  End Select

  FormatAccount     = strFmtAccount

End Function


Function FormatHost(strHostParm, strFDQN)
  Call DebugLog("FormatHost: " & strHostParm)
  Dim strAlias, strUserDomain, strUserDNSDomain

  strAlias          = strHostParm
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  If strUserDNSDomain <> "" Then
    strUserDNSDomain = "." & strUserDNSDomain
  End If

  If UCase(Right(strAlias, Len(strUserDNSDomain))) = UCase(strUserDNSDomain) Then
    strAlias        = Left(strAlias, Len(strAlias) - Len(strUserDNSDomain))
  End If 

  If strFDQN = "F" Then
    strAlias        = strAlias & strUserDNSDomain
  End If

  FormatHost       = strAlias

End Function


Function GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
  Call DebugLog("GetAccountAttr: " & strUserAccount & ", " & strUserAttr)
  Dim objACE, objAttr, objDACL, objField
  Dim strAccount,strAttrObject, strAttrItem, strAttrList, strAttrValue, strSearchAttr
 
  strAttrItem       = ""
  strAttrValue      = ""
  strAccount        = FormatAccount(strUserAccount)
  intIdx            = Instr(strAccount, "\")
  Select Case True
    Case intIdx > 0
      strAccount    = Mid(strAccount, intIdx  + 1)
    Case StrComp(strAccount, strServer, vbTextCompare) = 0
      strAccount    = strAccount & "$"
    Case StrComp(strAccount, strClusterName, vbTextCompare) = 0
      strAccount    = strAccount & "$"
  End Select

  Select Case True
    Case strUserAttr = "msDS-GroupMSAMembership"
      strSearchAttr = "msDS-ManagedPasswordId," & strUserAttr
    Case Else
      strSearchAttr = strUserAttr
  End Select

  On Error Resume Next 
  objADOCmd.CommandText          = "<LDAP://DC=" & Replace(strUserDNSDomain, ".", ",DC=") & ">;(&(sAMAccountName=" & strAccount & "));distinguishedName," & strSearchAttr
  Set objRecordSet  = objADOCmd.Execute

  On Error Goto 0
  Select Case True
    Case Not IsObject(objRecordset)
      ' Nothing
    Case objRecordset Is Nothing
      ' Nothing
    Case IsNull(objRecordset)
      ' Nothing
    Case objRecordset.RecordCount = 0 
      ' Nothing
    Case strUserAttr = "msDS-GroupMSAMembership"
      Select Case True
        Case IsNull(objRecordset.Fields(1).Value) ' ManagedPasswordId only present for gMSA
          ' Nothing
        Case IsNull(objRecordset.Fields(2).Value) ' Empty msDS-GroupMSAMembership
          strAttrValue = "> "
        Case Else
          strAttrValue   = "> " 
          Set objField   = GetObject("LDAP://" & objRecordset.Fields(0).Value)  
          Set objAttr    = objField.Get("msDS-GroupMSAMembership")
          Set objDACL    = objAttr.DiscretionaryAcl
          For Each objACE In objDACL
            strAttrItem  = objACE.Trustee
            If CheckGroup(strAttrItem, strUserDNSDomain) = True Then
              strAttrValue = strAttrValue & strAttrItem & " "
            End If
          Next
      End Select
    Case IsNull(objRecordset.Fields(1).Value)
      ' Nothing
    Case Instr(strUserAttr, "SID") > 0
      strAttrValue  = OctetToHexStr(objRecordset.Fields(1).Value)
      strAttrValue  = HexStrToSIDStr(strAttrValue)
    Case Instr(strUserAttr, "GUID") > 0
      strAttrValue  = OctetToHexStr(objRecordset.Fields(1).Value)
      strAttrValue  = HexStrToGUID(strAttrValue)
    Case strUserAttr = "memberOf"
      strAttrList   = ""
      For Each strAttrItem In objRecordset.Fields(1).Value
        strAttrList = strAttrList & Mid(strAttrItem, 4, Instr(strAttrItem, ",") - 4) & " "
      Next
      strAttrValue = RTrim(strAttrList)
    Case Else
      strAttrValue  = objRecordset.Fields(1).Value
  End Select

  err.Number        = 0
  GetAccountAttr    = strAttrValue

End Function


Function GetCertThumbprint(strCertName)
  Call DebugLog("GetCertThumbprint: " & strCertName)

  strCmd            = "POWERSHELL (Get-ChildItem -Path Cert:\LocalMachine\My ^| Where-Object {$_.FriendlyName -match '" & strCertName & "'}).Thumbprint"
  Set objExec       = objShell.Exec(strCmd)
  GetCertThumbprint = LCase(objExec.StdOut.ReadAll)

End Function


Function GetOUAttr(strOUPath, strUserDNSDomain, strOUAttr)
  Call DebugLog("GetOUAttr: " & strOUPath & ", " & strOUAttr)
  Dim objOU
  Dim arrOUPath
  Dim strAttrValue, strOUCName, strCName, strDelim

  Select Case True
    Case Instr(strOUPath, "/") > 0
      strDelim      = "/"
    Case Else
      strDelim      = "."
  End Select
  arrOUPath         = Split(Replace("OU=" & strOUPath, strDelim, ".OU="), ".")
  strOUCName        = Replace("DC=" & strUserDNSDomain, ".", ",DC=")
  For Each strCName In arrOUPath
    strOUCName      = strCName & "," & strOUCName
  Next
  Call DebugLog("OU CName: " & strOUCName)

  On Error Resume Next 
  Set objOU         = GetObject("LDAP://" & strOUCName)

  On Error Goto 0
  strAttrValue      = ""
  Select Case True
    Case Not IsObject(objOU)
      ' Nothing
    Case objOU Is Nothing
      ' Nothing
    Case IsNull(objOU)
      ' Nothing
    Case Instr(strOUAttr, "GUID") > 0
      strAttrValue  = objOU.Get(strOUAttr)
      strAttrValue  = OctetToHexStr(strAttrValue)
      strAttrValue  = HexStrToGUID(strAttrValue)
    Case Else
      strAttrValue  = objOU.Get(strOUAttr)
  End Select

  err.Number        = 0
  GetOUAttr         = strAttrValue

End Function


Private Function OctetToHexStr(strValue)
  Dim strHexStr
  Dim intIdx

  strHexStr         = ""
  For intIdx = 1 To Lenb(strValue)
    strHexStr       = strHexStr & Right("0" & Hex(Ascb(Midb(strValue, intIdx, 1))), 2)
  Next

  OctetToHexStr     = strHexStr

End Function


Private Function HexStrToGUID(strValue)
  Dim strGUID

  strGUID           = ""

  If Len(strValue) = 32 Then
    strGUID         = strGUID & Mid(strValue,  7, 2) & Mid(strValue,  5, 2) & Mid(strValue,  3, 2) & Mid(strValue,  1, 2) & "-"
    strGUID         = strGUID & Mid(strValue, 11, 2) & Mid(strValue,  9, 2) & "-"
    strGUID         = strGUID & Mid(strValue, 15, 2) & Mid(strValue, 13, 2) & "-"
    strGUID         = strGUID & Mid(strValue, 17, 2) & Mid(strValue, 19, 2) & "-"
    strGUID         = strGUID & Mid(strValue, 21)
  End If

  HexStrToGUID      = strGUID

End Function


Private Function HexStrToSIDStr(strValue)
  Dim arrSID
  Dim strSIDStr
  Dim intIdx, intUB, intWork

  intUB             = (Len(strValue) / 2) - 1
  ReDim arrSID(intUB)  
  For intIdx = 0 To intUB
    arrSID(intIdx)  = CInt("&H" & Mid(strValue, (intIdx * 2) + 1, 2))
  Next

  strSIDStr         = "S-" & arrSID(0) & "-" & arrSID(1) & "-" & arrSID(8)
  If intUB >= 15 Then
    intWork         = arrSID(15)
    intWork         = (intWork * 256) + arrSID(14)
    intWork         = (intWork * 256) + arrSID(13)
    intWork         = (intWork * 256) + arrSID(12)
    strSIDStr       = strSIDStr & "-" & CStr(intWork)
    If intUB >= 22 Then
      intWork       = arrSID(19)
      intWork       = (intWork * 256) + arrSID(18)
      intWork       = (intWork * 256) + arrSID(17)
      intWork       = (intWork * 256) + arrSID(16)
      strSIDStr     = strSIDStr & "-" & CStr(intWork)
      intWork       = arrSID(23)
      intWork       = (intWork * 256) + arrSID(22)
      intWork       = (intWork * 256) + arrSID(21)
      intWork       = (intWork * 256) + arrSID(20)
      strSIDStr     = strSIDStr & "-" & CStr(intWork)
    End If
    If intUB >= 25 Then
      intWork       = arrSID(25)
      intWork       = (intWork * 256) + arrSID(24)
      strSIDStr     = strSIDStr & "-" & CStr(intWork)
    End If
  End If

  HexStrToSIDStr    = strSIDStr

End Function


Sub ProcessUser(strLabel, strDescription, strProcess)
  Call SetProcessId(strLabel, strDescription)

  For Each objFolder In arrProfFolders
    Select Case True
      Case Not objFSO.FileExists (objFolder.Path & "\NTUSER.DAT")
        ' Nothing
      Case Else
        Call DebugLog("Account path: " & objFolder.Path)
        strCmd      = strProgReg & " LOAD ""HKLM\FBTempProf"" """ & objFolder.Path & "\NTUSER.DAT"""
        Call Util_RunExec(strCmd, "", strResponseYes, -1)
        Select Case True
          Case intErrSave = 0
            strCmd  = "Call " & strProcess & "(" & strHKLM & ", ""HKLM\"", ""FBTempProf"")"
            strDebugMsg1 = strCmd
            Execute strCmd
            strCmd  = strProgReg & " UNLOAD  ""HKLM\FBTempProf"""
            Call Util_RunExec(strCmd, "", strResponseYes, 1)
          Case intErrSave = 1
            Call DebugLog("Processing bypassed for " & objFolder.Path)
          Case intErrSave = 1332
            Call DebugLog("Processing bypassed for " & objFolder.Path)
          Case Else
            Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
        End Select
    End Select
  Next

  For Each strUser In arrProfUsers
    Select Case True
      Case Right(strUser, 8) = "_Classes"
        ' Nothing
      Case (Len(strUser) = 8) And (strUser <> ".DEFAULT")
        ' Nothing - Local system account
      Case Else
        Call DebugLog("Account SID: " & strUser)
        Execute "Call " & strProcess & "(" & strHKU & ", ""HKEY_USERS\"", """ & strUser & """)"
    End Select
  Next

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ResetDBAFilePerm(strFolder)
  Call DebugLog("ResetDBAFilePerm: " & strFolder)

  Call ResetFilePerm(strFolder, strGroupDBA)

  If strGroupDBANonSA <> "" Then
    Call ResetFilePerm(strFolder, strGroupDBANonSA)
  End If

  If strIsInstallDBA = "1" Then
    Call ResetFilePerm(strFolder, strUserAccount)
  End If

End Sub


Sub ResetFilePerm(strFolder, strAccount)
  Call DebugLog("ResetFilePerm: " & strAccount)

  strPath           = strFolder
  If Right(strPath, 1) = "\" Then
    strPath         = Left(strPath, Len(strPath) - 1)
  End If

  Select Case True
    Case strAccount = strGroupDBA
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strGroupDBA) & """:F"
      Call SetFilePerm(strCmd)
    Case strAccount = strGroupDBANonSA
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strGroupDBANonSA) & """:R"
      Call SetFilePerm(strCmd)
    Case Else 
      strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strAccount) & """:F"
      Call SetFilePerm(strCmd)
  End Select

End Sub


Sub SetDCOMSecurity(strAppId)
  Call DebugLog("SetDCOMSecurity: " & strAppId)
  Dim arrPermDCom
  Dim objHelper, objPermDCom
  Dim strDescription, strPermDCom, strSDDLDCom

  objWMIReg.GetBinaryValue strHKCR,strAppId,"LaunchPermission",arrPermDCom
  Select Case True
    Case IsNull(arrPermDCom) 
      Exit Sub
    Case strOSVersion < "6.0"
      Exit Sub
  End Select

  objWMIReg.GetStringValue strHKCR,strAppId,"",strDescription
  Call DebugLog(" " & strDescription & ", Appid: " & strAppId & ", Current Perm: " & strPermDCom)  

  strSDDLDCom       = "(A;;CCDCLCSWRP;;;" & strSIDDistComUsers & ")"
  strPath           = "winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\ROOT\cimv2:Win32_securityDescriptorHelper"
  Set objHelper     = GetObject(strPath)
  Call objHelper.BinarySDToSDDL(arrPermDCom, strPermDCom)
  intIdx            = Instr(strPermDCom, strSIDDistComUsers)
  If intIdx = 0 Then
    intIdx          = Instr(strPermDCom, "(A;;CCSW;;;BU)")
    If intIdx = 0 Then
      strPermDCom   = strPermDCom & strSDDLDCom 
    Else
      strPermDCom   = Left(strPermDCom, intIdx - 1) & strSDDLDCom & Mid(strPermDCom, intIdx)
    End If
    Call DebugLog("Update DCom security with " & strPermDCom)
    Call objHelper.SDDLToWin32SD(strPermDCom, objPermDCom)
    Call objHelper.Win32SDToBinarySD(objPermDCom, arrPermDCom)
    objWMIReg.SetBinaryValue strHKCR,strAppId,"LaunchPermission",arrPermDCom
  End If

  Set objHelper     = Nothing

End Sub


Sub SetFilePerm(strFolderPerm)
  Call DebugLog("SetFilePerm: " & strFolderPerm)
  Dim arrFolderPerm
  Dim intUBound, intIdx, intIdx2
  Dim strNTService, strShareDrive

  arrFolderPerm     = Split(strFolderPerm)
  intUBound         = UBound(arrFolderPerm)
  strNTService      = GetBuildfileValue("NTService")
  For intIdx = 0 To intUBound
    Select Case True
      Case Instr(arrFolderPerm(intIdx), """:") = 0 
        ' Nothing
      Case Instr(arrFolderPerm(intIdx), strNTService & "\") > 0 
        arrFolderPerm(intIdx) = ""
      Case Else
        For intIdx2 = intIdx + 1 To intUBound
          Select Case True
            Case Instr(arrFolderPerm(intIdx2), """:") = 0
              ' Nothing
            Case StrComp(Left(arrFolderPerm(intIdx), Instr(arrFolderPerm(intIdx), """:")), Left(arrFolderPerm(intIdx2), Instr(arrFolderPerm(intIdx2), """:")), vbTextCompare) = 0
              arrFolderPerm(intIdx) = ""
          End Select
        Next      
    End Select
  Next  
  strFolderPerm            = Join(arrFolderPerm, " ")

  intIdx2           = 0
  For intIdx = 0 To intUBound
    Select Case True
      Case Instr(arrFolderPerm(intIdx), """:") = 0 
        ' Nothing
      Case Else
        intIdx2     = 1
    End Select
  Next
  If intIdx2 = 0 Then
    Exit Sub
  End If

  strShareDrive     = ""
  If Instr(strFolderPerm, "\\") > 0 Then
    strShareDrive   = GetShareDrive(strFolderPerm)
  End If

  Call Util_RunExec(strProgCacls & " " & strFolderPerm, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 2
      ' Nothing
    Case intErrSave = 13
      ' Nothing
    Case intErrSave = 67     ' Network Name not found
      ' Nothing
    Case intErrSave = 1240   ' Not Authorized - Cannot put permission on remote share root
      ' Nothing
    Case intErrSave = 1332   ' Problem with security descriptor
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strFolderPerm)
  End Select
  Wscript.Sleep strWaitShort ' Allow time for CACLS processing to complete

  If strShareDrive <> "" Then
    Call Util_RunExec("NET USE " & strShareDrive & " /DELETE", "EOF", "", -1)
  End If

End Sub


Private Function GetShareDrive(strCmd)
  Call DebugLog("GetShareDrive: " & strCmd)
  Dim intIdx, intIdx1, intIdx2, intIdx3, intIdx4
  Dim strAlphabet, strDriveList, strShare, strShareDrive

  strAlphabet       = GetBuildfileValue("Alphabet")
  strDriveList      = GetBuildfileValue("DriveList")
  strShareDrive     = ""
  For intIdx = 3 To Len(strAlphabet)
    strDebugMsg1    = "Index " & CStr(intIdx)
    If Instr(strDriveList, Mid(strAlphabet, intIdx, 1)) = 0 Then
      strDebugMsg2    = "Drive Found"
      strShareDrive = Mid(strAlphabet, intIdx, 1) & ":"
      Exit For
    End If
  Next

  If strShareDrive <> "" Then
    intIdx          = Instr(strCmd, "\\")
    intIdx1         = Instr(intIdx  + 2, strCmd, "\")
    intIdx2         = Instr(intIdx1 + 1, strCmd, "\")
    intIdx3         = Instr(intIdx1 + 1, strCmd, """")
    If intIdx3 = 0 Then
      intIdx3       = Len(strCmd)
    End If
    intIdx4         = Min(intIdx2, intIdx3)
    strShare        = Mid(strCmd, intIdx, intIdx4 - intIdx)
    Call Util_RunExec("NET USE " & strShareDrive & " """ & strShare & """ /PERSISTENT:NO", "EOF", "", 0)
    strCmd          = Left(strCmd, intIdx - 1) & strShareDrive & Mid(strCmd, intIdx4)
    Wscript.Sleep strWaitShort
  End If

  GetShareDrive     = strShareDrive

End Function


Sub SetFWRule(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
  Call DebugLog("SetFWRule: " & strFWName & " for " & strFWPort)

  Select Case True
    Case Left(strOSVersion, 1) < "6"
      Call SetFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
    Case Else
      Call SetAdvFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
  End Select

End Sub


Private Sub SetFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
  Call DebugLog("SetFirewall:")
  Dim strRuleExist, strRuleType
  
  strRuleExist      = CheckFWName(strFWName)
  Select Case True
    Case strFWProgram <> ""
      strRuleType   = "ALLOWEDPROGRAM"
    Case Else
      strRuleType   = "PORTOPENING"
  End Select

  Select Case True
    Case strFirewallStatus <> "1"
      ' Nothing
    Case strRuleExist = False
      strCmd        = "NETSH FIREWALL ADD " & strRuleType & " NAME=""" & strFWName & """ "
      strCmd        = strCmd & "MODE=ENABLE SCOPE=ALL PROFILE=DOMAIN "
      If strFWType <> "" Then
        strCmd      = strCmd & "PROTOCOL=" & strFWType & " "
      End If
      If strFWPort <> "" Then
        strCmd      = strCmd & "PORT=" & Replace(strFWPort, " ", "") & " "
      End If
      If strFWProgram <> "" Then
        strCmd      = strCmd & "PROGRAM=""" & strFWProgram & """ "
      End If
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  If (strRuleExist = True) Or (strFWEnable = "Y") Then
    strCmd          = "NETSH FIREWALL SET NAME=""" & strFWName & """ "
    strCmd          = strCmd & "PROFILE=DOMAIN MODE=ENABLE  "
'    Call Util_RunExec(strCmd, "", strResponseYes, 0) verify syntax correct before enabling command
  End If

End Sub


Private Sub SetAdvFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
  Call DebugLog("SetAdvFirewall:")
  Dim strRuleExist

  strRuleExist      = CheckFWName(strFWName)

  If strRuleExist = False Then 
    strCmd          = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""" & strFWName & """ "
    strCmd          = strCmd & "ACTION=ALLOW PROFILE=DOMAIN "
    If strFWDesc <> "" Then
      strCmd        = strCmd & "DESCRIPTION=""" & strFWDesc & """ "
    End If
    If strFWType <> "" Then
      strCmd        = strCmd & "PROTOCOL=" & strFWType & " "
    End If
    If strFWDir <> "" Then
      strCmd        = strCmd & "DIR=" & strFWDir & " "
    End If
    If strFWPort <> "" Then
      strCmd        = strCmd & "LOCALPORT=" & Replace(strFWPort, " ", "") & " "
    End If
    If strFWProgram <> "" Then
      strCmd        = strCmd & "PROGRAM=""" & strFWProgram & """ "
    End If
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

  If (strRuleExist = True) Or (strFWEnable = "Y") Then
    strCmd          = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""" & strFWName & """ "
    strCmd          = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES "
    Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End If

End Sub


Private Function CheckFWName(strFWName)
  Call DebugLog("CheckFWName:")
  Dim objFWRule

  CheckFWName       = False
  For Each objFWRule In objFWRules
    If objFWRule.Name = strFWName Then
      CheckFWName   = True
      Exit For
    End If
  Next

End Function


Sub SetRegPerm(strRegParm, strName, strAccess)
  Call DebugLog("SetRegPerm: " & strRegParm & " for " & strName)
  ' Code based on example posted by ROHAM on www.tek-tips.com/viewthread.cfm?qid=1456390
  Dim objACE, objDACL, objSD
  Dim intBypass, intIdx
  Dim strACEAccessAllow, strACEFullControl, strACEPropogate, strACERead, strPath, strRegKey, strRegRoot, strRegType, strSDFormatIID, strSDPathRegistry, strTrusteeName

  intIdx            = Instr(strRegParm, "\")
  strRegType        = 0
  Select Case True
    Case UCase(Left(strRegParm, intIdx)) = "HKEY_CLASSES_ROOT\"
      strRegType    = &H80000000
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = Left(strRegParm, intIdx)
    Case UCase(Left(strRegParm, intIdx)) = "HKCR\"
      strRegType    = &H80000000
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_CLASSES_ROOT\"
    Case UCase(Left(strRegParm, intIdx)) = "HKEY_CURRENT_USER\"
      strRegType    = &H80000001
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = Left(strRegParm, intIdx)
    Case UCase(Left(strRegParm, intIdx)) = "HKCU\"
      strRegType    = &H80000001
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_CURRENT_USER\"
    Case UCase(Left(strRegParm, intIdx)) = "HKEY_LOCAL_MACHINE\"
      strRegType    = &H80000002
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = Left(strRegParm, intIdx)
    Case UCase(Left(strRegParm, intIdx)) = "HKLM\"
      strRegType    = &H80000002
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_LOCAL_MACHINE\"
    Case UCase(Left(strRegParm, intIdx)) = "HKEY_USERS\"
      strRegType    = &H80000003
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = Left(strRegParm, intIdx)
    Case UCase(Left(strRegParm, intIdx)) = "HKUS\"
      strRegType    = &H80000003
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_USERS\"
    Case UCase(Left(strRegParm, intIdx)) = "HKU\"
      strRegType    = &H80000003
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_USERS\"
    Case UCase(Left(strRegParm, intIdx)) = "HKEY_CURRENT_CONFIG\"
      strRegType    = &H80000005
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = Left(strRegParm, intIdx)
    Case UCase(Left(strRegParm, intIdx)) = "HKCC\"
      strRegType    = &H80000005
      strRegKey     = Mid(strRegParm, intIdx + 1)
      strRegRoot    = "HKEY_CURRENT_CONFIG\"
  End Select

  Select Case True
    Case strRegType = 0
      intBypass     = 1
    Case objWMIReg.EnumKey(strRegType, strRegKey, "") <> 0
      Select Case True
        Case strRegRoot <> "HKEY_USERS\"
          intBypass = 1
        Case Left(strRegKey, 9) = ".DEFAULT\"
          If Right(strRegKey, 1) <> "\" Then
            strRegKey = strRegKey & "\"
          End If
          Call Util_RegWrite(strRegRoot & strRegKey, "", "REG_SZ")
          intBypass = 0
        Case Else
          intBypass = 1
      End Select
    Case GetBuildfileValue("CheckRegPerm") <> "OK"
      intBypass     = 1
    Case Else
      intBypass     = 0
  End Select
  If intBypass = 1 Then
    Call DebugLog("SetRegPerm bypassed")
    Exit Sub
  End If

  strACEAccessAllow = 0
  strACEFullControl = &h10000000
  strACEPropogate   = &h2
  strACERead        = &h80000000
  strSDFormatIID    = 1
  strSDPathRegistry = 3
  If Right(strRegKey, 1) = "\" Then
    strRegKey       = Left(strRegKey, Len(strRegKey) - 1)
  End If

  strTrusteeName    = FormatAccount(strName)
  Set objSD         = objSDUtil.GetSecurityDescriptor(strRegRoot & strRegKey, strSDPathRegistry, strSDFormatIID)
  objSD.Owner       = strLocalAdmin
  Set objDACL       = objSD.DiscretionaryAcl
  Set objACE        = CreateObject("AccessControlEntry")
  objACE.Trustee    = strTrusteeName
  Select Case True
    Case strAccess = "F"
      objACE.AccessMask = strACEFullControl
    Case Else
      objACE.AccessMask = strACERead
  End Select
  objACE.ACEType    = strACEAccessAllow
  objACE.ACEFlags   = strACEPropogate
  objDACL.AddAce objACE

  objSDUtil.SetSecurityDescriptor strRegRoot & strRegKey, strSDPathRegistry, objSD, strSDFormatIID

  Set objACE        = Nothing
  Set objDACL       = Nothing
  Set objSD         = Nothing

End Sub


End Class


Sub BackupDBMasterKey(strDB, strPassword)
  Call FBManageSecurity.BackupDBMasterKey(strDB, strPassword)
End Sub

Function FormatAccount(strAccount)
  FormatAccount     = FBManageSecurity.FormatAccount(strAccount)
End Function

Function FormatHost(strHostParm, strFDQN)
  FormatHost       = FBManageSecurity.FormatHost(strHostParm, strFDQN)
End Function

Function GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
  GetAccountAttr    = FBManageSecurity.GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
End Function

Function GetCertThumbprint(strCertName)
  GetCertThumbprint = FBManageSecurity.GetCertThumbprint(strCertName)
End Function

Function GetOUAttr(strOUPath, strUserDNSDomain, strOUAttr)
  GetOUAttr         = FBManageSecurity.GetOUAttr(strOUPath, strUserDNSDomain, strOUAttr)
End Function

Sub ProcessUser(strLabel, strDescription, strProcess)
  Call FBManageSecurity.ProcessUser(strLabel, strDescription, strProcess)
End Sub

Sub ResetDBAFilePerm(strFolder)
  Call FBManageSecurity.ResetDBAFilePerm(strFolder)
End Sub

Sub ResetFilePerm(strFolder, strAccount)
  Call FBManageSecurity.ResetFilePerm(strFolder, strAccount)
End Sub

Sub RunCacls(strFolderPerm)
  Call FBManageSecurity.SetFilePerm(strFolderPerm)
End Sub

Sub SetDCOMSecurity(strAppId)
  Call FBManageSecurity.SetDCOMSecurity(strAppId)
End Sub

Sub SetFilePerm(strFolderPerm)
  Call FBManageSecurity.SetFilePerm(strFolderPerm)
End Sub

Sub SetFWRule(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
  Call FBManageSecurity.SetFWRule(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWDesc, strFWEnable)
End Sub

Sub SetRegPerm(strRegParm, strName, strAccess)
  Call FBManageSecurity.SetRegPerm(strRegParm, strName, strAccess)
End Sub