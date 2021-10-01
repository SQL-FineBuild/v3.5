'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBUtils.vbs  
'  Copyright FineBuild Team © 2017 - 2021.  Distributed under Ms-Pl License
'
'  Purpose:      Miscellaneous Utilities 
'
'  Author:       Ed Vassie
'
'  Date:         05 Jul 2017
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     05 Jul 2017  Initial version

'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim FBUtils: Set FBUtils = New FBUtilsClass
Dim intErrSave
Dim strErrSave, strResponseYes, strResponseNo

Class FBUtilsClass

Dim objAutoUpdate, objExec, objFile, objFSO, objShell, objSQL, objSQLData, objWMI, objWMIReg
Dim colPrcEnvVars
Dim intIdx
Dim strCmd, strCmdPS, strCmdSQL
Dim strOSVersion
Dim strPath, strPathCmdSQL, strPathTemp, strPathTools, strProgCacls, strRegTools
Dim strServer, strServInst, strType, strSQLVersion, strSQLVersionNum, strWaitShort


Private Sub Class_Initialize
  Call DebugLog("FBUtils Class_Initialize:")

  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objShell      = CreateObject("Wscript.Shell")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")

  intErrSave        = 0
  strErrSave        = ""

  If strProcessIdCode <> "FBCV" Then
    strCmdPS        = GetBuildfileValue("CmdPS")
    strCmdSQL       = GetBuildfileValue("CmdSQL")
    strOSVersion    = GetBuildfileValue("OSVersion")
    strPathTemp     = GetBuildfileValue("PathTemp")
    strProgCacls    = GetBuildfileValue("ProgCacls")
    strServer       = GetBuildfileValue("AuditServer")
    strServInst     = GetBuildfileValue("ServInst")
    strType         = GetBuildfileValue("Type")
    strResponseNo   = GetBuildfileValue("ResponseNo")
    strResponseYes  = GetBuildfileValue("ResponseYes")
    strWaitShort    = GetBuildfileValue("WaitShort")
    Call SetHKLMSQL()
  End If

End Sub


Sub CopyFile(strSource, strTarget)
  Call DebugLog("CopyFile: " & strSource & " to " & strTarget)
  Dim strTgtFolder

  Select Case True
    Case Right(strTarget, 1) = "\"
      strTgtFolder  = strTarget
    Case Else
      strTgtFolder  = strTarget & "\"
  End Select

  If objFSO.FileExists(strSource) Then
    Set objFile     = objFSO.GetFile(strSource)
    strPath         = strTgtFolder & objFile.Name
    If Not objFSO.FileExists(strPath) Then
      objFile.Copy strPath, True
    End If
  End If

End Sub


Sub DeleteFile(strFile)
  Call DebugLog("DeleteFile: " & strFile)

  If objFSO.FileExists(strFile) Then
    Call objFSO.DeleteFile(strFile, True)
    Wscript.Sleep strWaitShort
  End If

End Sub


Sub DeleteFolder(strFolder)
  Call DebugLog("DeleteFolder: " & strFolder)
  Dim strPath

  Select Case True
    Case strFolder = ""
      Exit Sub
    Case Right(strFolder, 1) = "\" 
      strPath       = Left(strFolder, Len(strFolder) - 1)
    Case Else
      strPath       = strFolder
  End Select

  Select Case True
    Case strPath = strPathTemp
      ' Nothing
    Case Not objFSO.FolderExists(strPath)
      ' Nothing
    Case Else
      objFSO.DeleteFolder strPath, 1
      Wscript.Sleep strWaitShort
  End Select

End Sub


Function FormatFolder(strFolder)
  Call DebugLog("FormatFolder: " & strFolder)
  Dim strFBLocal, strFBRemote, strFolderPath

  strFBLocal        = GetBuildfileValue("FBPathLocal")
  strFBRemote       = GetBuildfileValue("FBPathRemote")

  Select Case True
    Case strFolder = ""
      strFolderPath = ""
    Case Mid(strFolder, 2, 1) = ":"
      strFolderPath = strFolder
    Case Left(strFolder, 2) = "\\"
      strFolderPath = strFolder
    Case Else
      strFolderPath = GetBuildfileValue(strFolder)
  End Select
  
  Select Case True
    Case strFBLocal = strFBRemote
      ' Nothing
    Case UCase(Left(strFolderPath, Len(strFBRemote))) = strFBRemote
      strFolderPath = strFBLocal & Mid(strFolderPath, Len(strFBRemote) + 1)
  End Select

  Select Case True
    Case strFolderPath = ""
      ' Nothing
    Case Right(strFolderPath, 1) = "\"
      ' Nothing
    Case Else
      strFolderPath = strFolderPath & "\"
  End Select

  FormatFolder      = strFolderPath  
 
End Function


Function FormatFolderURI(strFolder)
  Call DebugLog("FormatFolderURI: " & strFolder)
  Dim strFolderPath

  strFolderPath     = FormatFolder(strFolder)

  strFolderPath     = Replace(Replace(strFolderPath, "\", "/"), "%", "%25")
  strFolderPath     = Replace(Replace(Replace(Replace(Replace(Replace(strFolderPath, " ", "%20"), "#", "%23"), "$", "%24"), "?", "%3F"), "{", "%7B"), "}", "%7D")
  strFolderPath     = "file:///" & strFolderPath

  FormatFolderURI   = strFolderPath

End Function


Function FormatServer(strServer, strProtocol)
  Dim strServerWork

  strServerWork     = strServer
  If strProtocol <> "" Then
    strServerWork   = strProtocol & "://" & FormatHost(strServerWork, "F")
  End If

  FormatServer      = strServerWork

End Function


Function Max(intA, intB)

  Select Case True
    Case CLng(intA) > CLng(intB)
      Max           = intA
    Case Else
      Max           = intB
  End Select

End Function


Function Min(intA, intB)

  Select Case True
    Case CLng(intA) > CLng(intB)
      Min           = intB
    Case Else
      Min           = intA
  End Select

End Function


Function GetCmdSQL()
  Call DebugLog("GetCmdSQL:")

  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strRegTools       = strHKLMSQL & strSQLVersionNum & "\Tools\ClientSetup\"

  Select Case True
    Case strSQLVersion = "SQL2005"
      objWMIReg.GetStringValue strHKLM,Mid(strRegTools, 6),"Path",strPathTools
    Case strSQLVersion <= "SQL2012"
      objWMIReg.GetStringValue strHKLM,Mid(strRegTools, 6),"SQLPath",strPathTools
    Case Else
      objWMIReg.GetStringValue strHKLM,Mid(strRegTools, 6),"ODBCToolsPath",strPathTools
  End Select
  If IsNull(strPathTools) Then
    strPathTools    = ""
  End If

  Select Case True
    Case strPathTools = ""
      strPathCmdSQL = ""
    Case strSQLVersion = "SQL2005"
      strPathCmdSQL = strPathTools & "SQLCMD.EXE"
    Case strSQLVersion <= "SQL2012"
      strPathCmdSQL = strPathTools & "\Binn\SQLCMD.EXE"
    Case Else
      strPathCmdSQL = strPathTools & "SQLCMD.EXE"
  End Select

  Select Case True
    Case strPathcmdSQL = ""
      strCmdSQL     = ""
    Case Else
      strCmdSQL     = """" & strPathCmdSQL & """ -S """ & strServInst & """ -E -b -e -m-1 "
  End Select

  Call SetBuildfileValue("CmdSQL",     strCmdSQL)
  Call SetBuildfileValue("PathCmdSQL", strPathCmdSQL)
  Call SetBuildfileValue("RegTools",   strRegTools)

  GetCmdSQL         = strCmdSQL

End Function


Function GetPSData(strCmd)
  Call DebugLog("GetPSData: " & strCmd)
  Dim strPSData

  Set objExec       = objShell.Exec("POWERSHELL " & strCmd)
  objExec.StdIn.Close
  strPSData         = objExec.StdOut.ReadAll
  strPSData         = Left(strPSData, Len(strPSData) - 2) & Replace(Replace(Right(strPSData, 2), Chr(10), ""), Chr(13), "")
  GetPSData         = strPSData

End Function


Function GetSQLConnection(strServInst)
  Call DebugLog("GetSQLConnection: " & strServInst)
  
  Set objSQL        = CreateObject("ADODB.Connection")
  objSQL.Provider   = "SQLOLEDB"
  objSQL.ConnectionString = "Driver={SQL Server};Server=" & strServInst & ";Database=master;Trusted_Connection=Yes;"
  strDebugMsg1      = objSQL.ConnectionString
  objSQL.Open 

  Set objSQLData    = CreateObject("ADODB.Recordset")
  Set GetSQLConnection = objSQL

End Function


Function GetSQLData(strCmd, strCmdItem)
  Call DebugLog("GetSQLData: " & strCmd & ": " & strCmdItem)

  Set objSQLData    = objSQL.Execute(strCmd)
  objSQLData.MoveFirst
  GetSQLData        = objSQLData.Fields(strCmdItem)

End Function


Sub SetupFolder(objParm)
  Dim objParmParent
  Dim strPath, strPathParent, strSecurity, strShare

  Select Case True
    Case IsObject(objParm) = 0
      strPath       = objParm
      strSecurity   = ""
    Case Else
      strPath       = GetXMLParm(objParm, "Folder",    "")
      strSecurity   = GetXMLParm(objParm, "Security",  "")
  End Select
  Call DebugLog("SetupFolder: " & strPath)

  Select Case True
    Case strPath = ""
      Exit Sub
    Case Right(strPath, 1) = "\" 
      strPath       = Left(strPath, Len(strPath) - 1)
  End Select
  Select Case True
    Case Left(strPath, 2) <> "\\"
      ' Nothing
    Case Instr(3, strPath, "\") = 0
      strShare      = strPath
    Case Else
      strShare      = Left(strPath, Instr(3, strPath, "\") - 1)
  End Select

  strPathParent     = Left(strPath, InstrRev(strPath, "\") - 1)
  strDebugMsg1      = "PathParent: " & strPathParent

  Select Case True
    Case objFSO.FolderExists(strPath)
      ' Nothing
    Case Not objFSO.FolderExists(strPathParent)
      Call SetXMLParm(objParmParent, "Folder",   strPathParent)
      Call SetXMLParm(objParmParent, "Security", strSecurity)
      Call SetupFolder(objParmParent)
      objFSO.CreateFolder(strPath)
      Wscript.Sleep strWaitShort
    Case Else
      objFSO.CreateFolder(strPath)
      Wscript.Sleep strWaitShort
  End Select

End Sub


Private Sub SetHKLMSQL()

  Select Case True
    Case GetBuildfileValue("WOWX86") = "TRUE"
      strHKLMSQL    = "HKLM\SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\"
    Case Else
      strHKLMSQL    = "HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\"
  End Select

  Call SetBuildfileValue("HKLMSQL", strHKLMSQL)

End Sub


Sub SetParam(strParamName, strParam, strNewValue, strMessage, ByRef strList)
  Call DebugLog("SetParam: " & strParamName & " to " & strNewValue)
  Dim strBuildValue

  strBuildValue     = GetBuildfileValue(strParamName)
  Select Case True
    Case strParam = strNewValue
      ' Nothing
    Case strParam = "N/A"
      ' Nothing
    Case (strParam = "NO") And (strNewValue = "N/A")
      strParam      = strNewValue
    Case (strParam = "") And (strNewValue = "YES") And (strMessage = "")
      strParam      = strNewValue
      strList       = strList & " " & strParamName
    Case strParam = ""
      strParam      = strNewValue
    Case strBuildValue = strNewValue
      strParam      = strNewValue
    Case Left(strType, 7) = "REFRESH"
      strParam      = strNewValue
    Case strMessage <> ""
      strParam      = strNewValue
      Call SetBuildMessage(strMsgInfo, Left("/" & strParamName & ":" & Space(24), Max(Len(strParamName) + 2, 24)) & " set to " & strNewValue & ": " & strMessage)
    Case strList <> ""
      strParam      = strNewValue
      strList       = strList & " " & strParamName
    Case Else
      strParam      = strNewValue
  End Select

End Sub


Sub SetUpdate(strOnOff)
  Call DebugLog("SetUpdate: messages " & strOnOff)
  On Error Resume Next

  Select Case True
    Case strOnOff <> "ON"
      ' Nothing
    Case strOSVersion > "6.2"
      colPrcEnvVars("SEE_MASK_NOZONECHECKS") = 1    ' Prevent Security Warning message hanging quiet install
      Call Util_RunExec("NET STOP wuauserv", "", "", 2)
    Case Else
      colPrcEnvVars("SEE_MASK_NOZONECHECKS") = 1    ' Prevent Security Warning message hanging quiet install
      err.Number    = objAutoUpdate.Pause()         ' Prevent Windows Update service triggering a reboot prompt
  End Select

  Select Case True
    Case strOnOff = "ON"
      ' Nothing
    Case strOSVersion > "6.2"
      colPrcEnvVars.Remove("SEE_MASK_NOZONECHECKS") ' Allow Security Warning messages
    Case Else
      colPrcEnvVars.Remove("SEE_MASK_NOZONECHECKS") ' Allow Security Warning messages
      err.Number    =  objAutoUpdate.Resume()       ' Resume normal Window Update Service prompts
  End Select

  Select Case True
    Case err.Number = 0
      ' No action
    Case err.Number < 0
      Call SetBuildMessage(strMsgWarning, "Error " & Cstr(err.Number) & " returned by Windows Update Service when setting service to " & strOnOff)
      err.Clear
    Case Else
      Call SetBuildMessage(strMsgError,   "Error " & Cstr(err.Number) & " returned by Windows Update configuration " & strOnOff)
  End Select

End Sub


Function GetXMLParm(objParm, strParm, strDefault)
  Dim strValue

  Select Case True
    Case IsObject(objParm) = 0
      strValue      = strDefault
    Case IsNull(objParm.documentElement.getAttribute(strParm))
      strValue      = strDefault
    Case Else
      strValue      = objParm.documentElement.getAttribute(strParm)
  End Select

  GetXMLParm       = strValue

End Function


Sub SetXMLParm(objParm, strParm, strValue)
  Dim objAttribute

  If IsObject(objParm) = 0 Then
    Set objParm     = CreateObject("MSXML2.DomDocument")
    Set objParm.documentElement = objParm.createElement("ROOT")
  End If

  Select Case True
    Case Not IsNull(objParm.documentElement.getAttribute(strParm))
      objParm.documentElement.setAttribute strParm, strValue
    Case Else
      Set objAttribute  = objParm.createAttribute(strParm)
      objAttribute.Text = strValue
      objParm.documentElement.Attributes.setNamedItem objAttribute
  End Select

End Sub


Sub SetXMLConfigValue(objConfig, strNode, strAttr, strValue, strType)
  Call DebugLog("SetXMLConfigValue: " & strAttr & ", " & strValue)
  Dim objNode, objAttr
  Dim strPrefix

  Select Case True
    Case strNode = ""
      ' Nothing
    Case IsNull(objConfig.documentElement.selectSingleNode(strNode))
      Set objNode   = objConfig.createElement(strNode)
      objConfig.appendChild objNode
    Case Else
      Set objNode   = objConfig.documentElement.selectSingleNode(strNode)
  End Select

  Select Case True
    Case strType <> "A"
      ' Nothing
    Case IsNull(objNode.getAttribute(strAttr))
      Set objAttr   = objConfig.createAttribute(strAttr)
      objAttr.Text  = strValue
      objNode.Attributes.setNamedItem objAttr
    Case Else
      objNode.setAttribute strAttr, strValue
  End Select

  Select Case True
    Case strType = "A"
      ' Nothing
    Case strNode = ""
      ' Nothing
    Case objNode.selectSingleNode(strAttr) Is Nothing
      Set objAttr   = objConfig.createElement(strAttr)
      objNode.appendChild objAttr
      If strValue <> "" Then
        objAttr.Text = strValue
     End If
    Case (objNode.selectSingleNode(strAttr).Text = "") And (strValue = "")
     ' Nothing
    Case Else
      Set objAttr   = objNode.selectSingleNode(strAttr)
      objAttr.Text  = strValue
  End Select

  Select Case True
    Case strType = "A"
      ' Nothing
    Case strNode <> ""
      ' Nothing
    Case objConfig.documentElement.selectSingleNode("//" & strAttr) Is Nothing
      Set objAttr   = objConfig.createElement(strAttr)
      objConfig.appendChild objAttr
      If strValue <> "" Then
        objAttr.Text = strValue
     End If
    Case (objConfig.documentElement.selectSingleNode("//" & strAttr).Text = "") And (strValue = "")
     ' Nothing
    Case Else
      Set objAttr   = objConfig.documentElement.selectSingleNode("//" & strAttr)
      objAttr.Text  = strValue
  End Select

End Sub


Sub Util_RegWrite(strRegKey, strRegValue, strRegType)
  Call DebugLog("Util_RegWrite: " & strRegKey)

  err.Number        = objShell.RegWrite(strRegKey, strRegValue, strRegType)
  intErrSave        = err.Number
  strErrSave        = err.Description
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by Write " & strRegKey)
  End Select

End Sub


Sub Util_RunCmdAsync(strCmd, strOK)
  Call DebugLog("Util_RunCmdAsync: " & strCmd)
  On Error Resume Next

  err.Number        = objShell.Run(strCmd,7,False)
  intErrSave        = err.Number
  strErrSave        = err.Description
  On Error Goto 0
  Select Case True
    Case strOK      = -1
      Call DebugLog("Command ended with code: " & intErrSave)
    Case intErrSave = 0 
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

  err.Clear

End Sub


Sub Util_ExecSQL(strCmd, strSQL, strOK)
  Call DebugLog("Util_ExecSQL: " & strSQL)
  Dim strSQLCmd

  strSQLCmd         = strCmd & " " & strSQL
  Call Util_RunExec(strSQLCmd, "EOF", strResponseYes, strOK)

End Sub


Sub Util_RunExec(strCmd, strMessage, strResponse, strOK)
  Call DebugLog("Util_RunExec: " & strCmd)
  Dim objCmd
  Dim intCount, intEOS
  Dim strBox1, strBox2, strStdOut

  On Error Resume Next
  err.Clear
  strBox1           = "[" & strResponseYes & "/" & strResponseNo & "]"
  strBox2           = "(" & strResponseYes & "/" & strResponseNo & ")?"
  Set objCmd        = objShell.Exec(strCmd)
  Select Case True
    Case Not IsObject(objCmd)
      intErrSave    = 8
      strErrSave    = "Command not recognised"
    Case Else
      Select Case True
        Case strMessage = "EOF"
          objCmd.StdIn.Close
        Case Left(strCmd, Len(strCmdPS)) = strCmdPS
          objCmd.StdIn.Close
        Case Left(strCmd, Len(strProgCacls) + 1) = strProgCacls & " "
          objCmd.StdIn.Write strResponse & vbCrLf
          objCmd.StdIn.Close
      End Select
      intEOS        = objCmd.StdOut.AtEndOfStream
      While Not intEOS
        strStdOut   = objCmd.StdOut.ReadLine()
        intEOS      = objCmd.StdOut.AtEndOfStream
        Select Case True
          Case objCmd.Status <> 0
            intEOS  = True
          Case Right(strStdOut, Len(strBox1)) = strBox1
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Right(strStdOut, Len(strBox2)) = strBox2
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Left(strStdOut, Len(strAnyKey)) = strAnyKey
            objCmd.StdIn.Write strResponse & vbCrLf
          Case strMessage = ""
            ' Nothing
          Case Right(strStdOut, Len(strMessage)) = strMessage
            objCmd.StdIn.Write strResponse & vbCrLf
        End Select
      Wend
      objCmd.StdIn.Close
      intCount      = 0
      intEOS        = objCmd.Status
      While intEOS = 0
        Wscript.Sleep strWaitShort
        intCount    = intCount + 1
        intEOS      = objCmd.Status
        If intCount > 10 Then
          intEOS    = intCount
        End If
      WEnd
      intErrsave    = objCmd.ExitCode
      Select Case True
        Case Left(strCmd, Len(strProgCacls) + 1) = strProgCacls & " "
          strErrSave = err.Description
        Case Else
          strErrSave = objCmd.StdErr.ReadAll()
      End Select
  End Select

  On Error Goto 0
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case Instr(" " & strOK & " ", " " & CStr(intErrSave) & " ") > 0
      ' Nothing
    Case strOK      = -1
      Call DebugLog("Command ended with code: " & intErrSave)
    Case intErrSave = 3010
      strReboot     = "Pending"
      Call SetBuildfileValue("RebootStatus", strReboot)
    Case intErrSave = -2067723326
      strReboot     = "Pending"
      Call SetBuildfileValue("RebootStatus", strReboot) 
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

  err.Clear

End Sub


End Class


Sub CopyFile(strSource, strTarget)
  Call FBUtils.CopyFile(strSource, strTarget)
End Sub

Sub DeleteFile(strFile)
  Call FBUtils.DeleteFile(strFile)
End Sub

Sub DeleteFolder(strFolder)
  Call FBUtils.DeleteFolder(strFolder)
End Sub

Function FormatFolder(strFolder)
  FormatFolder      = FBUtils.FormatFolder(strFolder)
End Function

Function FormatFolderURI(strFolder)
  FormatFolderURI   = FBUtils.FormatFolderURI(strFolder)
End Function

Function FormatServer(strServer, strProtocol)
  FormatServer      = FBUtils.FormatServer(strServer, strProtocol)
End Function

Function GetPSData(strCmd)
  GetPSData       = FBUtils.GetPSData(strCmd)
End Function

Function GetCmdSQL()
  GetCmdSQL         = FBUtils.GetCmdSQL()
End Function

Function GetSQLConnection(strServInst)
  Set GetSQLConnection = FBUtils.GetSQLConnection(strServInst)
End Function

Function GetSQLData(strCmd, strCmdItem)
  GetSQLData        = FBUtils.GetSQLData(strCmd, strCmdItem)
End Function

Function Max(intA, intB)
  Max               = FBUtils.Max(intA, intB)
End Function

Function Min(intA, intB)
  Min               = FBUtils.Min(intA, intB)
End Function

Sub SetupFolder(strFolder)
  Call FBUtils.SetupFolder(strFolder)
End Sub

Sub SetUpdate(strOnOff)
  Call FBUtils.SetUpdate(strOnOff)
End Sub

Function GetXMLParm(objParm, strParm, strDefault)
  GetXMLParm        = FBUtils.GetXMLParm(objParm, strParm, strDefault)
End Function

Sub SetXMLParm(objParm, strParm, strValue)
  Call FBUtils.SetXMLParm(objParm, strParm, strValue)
End Sub

Sub SetXMLConfigValue(objConfig, strNode, strAttr, strValue, strType)
  Call FBUtils.SetXMLConfigValue(objConfig, strNode, strAttr, strValue, strType)
End Sub

Sub SetParam(strParamName, strParam, strNewValue, strMessage, ByRef strList)
  Call FBUtils.SetParam(strParamName, strParam, strNewValue, strMessage, strList)
End Sub

Sub Util_RegWrite(strRegKey, strRegValue, strRegType)
  Call FBUtils.Util_RegWrite(strRegKey, strRegValue, strRegType)
End Sub

Sub Util_RunCmdAsync(strCmd, strOK)
  Call FBUtils.Util_RunCmdAsync(strCmd, strOK)
End Sub

Sub Util_ExecSQL(strCmd, strSQL, strOK)
  Call FBUtils.Util_ExecSQL(strCmd, strSQL, strOK)
End Sub

Sub Util_RunExec(strCmd, strMessage, strResponse, strOK)
  Call FBUtils.Util_RunExec(strCmd, strMessage, strResponse, strOK)
End Sub