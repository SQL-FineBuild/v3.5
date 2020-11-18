'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBManageAccount.vbs  
'  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
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
Dim FBManageAccount: Set FBManageAccount = New FBManageAccountClass

Class FBManageAccountClass
Dim objADOCmd, objADOConn, objFolder, objFSO, objSDUtil, objWMIReg
Dim arrProfFolders, arrProfUsers
Dim intIdx, intBuiltinDomLen, intNTAuthLen, intServerLen
Dim strBuiltinDom, strClusterName, strCmd, strHKLM, strHKU, strLocalAdmin, strNTAuth, strProfDir, strProgReg, strServer, strUser


Private Sub Class_Initialize
  Call DebugLog("FBManageAccount Class_Initialize:")

  Set objADOConn    = CreateObject("ADODB.Connection")
  Set objADOCmd     = CreateObject("ADODB.Command")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objSDUtil     = CreateObject("ADsSecurityUtility")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strHKLM           = &H80000002
  strHKU            = &H80000003
  strBuiltinDom     = GetBuildfileValue("BuiltinDom")
  strClusterName    = GetBuildfileValue("ClusterName")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strNTAuth         = GetBuildfileValue("NTAuth")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProgReg        = GetBuildfileValue("ProgReg")
  strServer         = GetBuildfileValue("AuditServer")

  Set arrProfFolders  = objFSO.GetFolder(strProfDir).SubFolders
  objWMIReg.EnumKey strHKU, "", arrProfUsers

  objADOConn.Provider            = "ADsDSOObject"
  objADOConn.Open "ADs Provider"
  Set objADOCmd.ActiveConnection = objADOConn

  intBuiltinDomLen  = Len(strBuiltinDom) + 1
  intNTAuthLen      = Len(strNTAuth) + 1
  intServerLen      = Len(strServer) + 1

End Sub


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


Function GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
  Call DebugLog("GetAccountAttr: " & strUserAccount & ", " & strUserAttr)
  Dim objACE, objAttr, objDACL, objField, objRecordSet
  Dim strAccount,strAttrObject, strAttrItem, strAttrList, strAttrValue
 
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

  On Error Resume Next 
  objADOCmd.CommandText          = "<LDAP://DC=" & Replace(strUserDNSDomain, ".", ",DC=") & ">;(&(sAMAccountName=" & strAccount & "));distinguishedName," & strUserAttr
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
    Case IsNull(objRecordset.Fields(1).Value)
      ' Nothing
    Case strUserAttr = "msDS-GroupMSAMembership"
      Set objField  = GetObject("LDAP://" & objRecordset.Fields(0).Value)
      Set objAttr   = objField.Get("msDS-GroupMSAMembership")
      Set objDACL   = objAttr.DiscretionaryAcl
      strAttrValue  = "> "
      For Each objACE In objDACL
        strAttrItem  = objACE.Trustee
        strAttrValue = strAttrValue & strAttrItem & " "
      Next
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


Function FormatAccount(strAccount)
  FormatAccount     = FBManageAccount.FormatAccount(strAccount)
End Function

Function GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
  GetAccountAttr    = FBManageAccount.GetAccountAttr(strUserAccount, strUserDNSDomain, strUserAttr)
End Function

Function GetOUAttr(strOUPath, strUserDNSDomain, strOUAttr)
  GetOUAttr         = FBManageAccount.GetOUAttr(strOUPath, strUserDNSDomain, strOUAttr)
End Function

Sub ProcessUser(strLabel, strDescription, strProcess)
  Call FBManageAccount.ProcessUser(strLabel, strDescription, strProcess)
End Sub

Sub SetRegPerm(strRegParm, strName, strAccess)
  Call FBManageAccount.SetRegPerm(strRegParm, strName, strAccess)
End Sub