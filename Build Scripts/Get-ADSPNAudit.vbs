''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Get-ADSPNAudit.vbs  
'  Copyright FineBuild Team © 2021.  Distributed under Ms-Pl License
'
'  Purpose:      Displays SPN and AllowedToDelegateTo information for AD accounts
'
'  Author:       Ed Vassie
'
'  Date:         December 2021
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     10 Dec 2021  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
On Error Goto 0

Dim objCommand, objConnection, objNetwork, objRecordSet, objRootDSE
Dim strDomainDN

Call Init()
Call Process()
Call Terminate()

Sub Init()

  Set objNetwork    = CreateObject("WScript.Network")
  Set objRootDSE    = GetObject ("LDAP://" & objNetwork.UserDomain & "/RootDSE")
  strDomainDN       = objRootDSE.Get("DefaultNamingContext")

  Set objConnection      = CreateObject("ADODB.Connection")
  objConnection.Provider = "ADsDSOObject"
  objConnection.Open "Active Directory Provider"

  Set objCommand    = CreateObject("ADODB.Command")
  objCommand.ActiveConnection = objConnection
  objCommand.Properties("Searchscope")   = 2 ' SUBTREE
  objCommand.Properties("Page Size")     = 250
  objCommand.Properties("Timeout")       = 30
  objCommand.Properties("Cache Results") = False
  objCommand.Properties("Sort on")       = "Name"
  objCommand.CommandText = "SELECT ADsPath FROM 'LDAP://" & strDomainDN & "'"
  Set objRecordSet  = objCommand.Execute

  wscript.echo "-- SPN Audit Report --"

End Sub


Sub Process()

  On Error Resume Next

  Do While Not objRecordSet.EOF
     If objRecordSet.Fields("Name") <> "" Then
       Call ProcessAccount(objRecordSet.Fields("ADsPath").Value)
     End If
     objRecordSet.MoveNext
  Loop

End Sub


Sub ProcessAccount(strADsPath)
  Dim objAccount, objACE, objAttr, objDACL
  Dim strAttr, strMsg

  On Error Resume Next

  Set objAccount    = GetObject(strADsPath)
  strMsg            = Mid(objAccount.Name, 4)
  If strMsg = "" Then
    Exit Sub
  End If

  Select Case True
    Case IsNull(objAccount.Get("msDS-ManagedPasswordId"))
      ' Nothing, Account is not a gMSA
    Case IsNull(objAccount.Get("msDS-GroupMSAMembership"))
      strMsg        = strMsg & vbCrLf & "  WARNING: No Group details for gMSA Account"
    Case Else
      strMsg        = strMsg & vbCrLf & "  gMSA Group Details:"
      Set objAttr   = objAccount.Get("msDS-GroupMSAMembership")
      Set objDACL   = objAttr.DiscretionaryAcl
      For Each objACE In objDACL
        strMsg      = strMsg & vbCRLF & "    " & objACE.Trustee
      Next
  End Select

  Select Case True
    Case IsNull(objAccount.Get("servicePrincipalName"))
      ' Nothing, no SPN definitions for Account
    Case Else
      strMsg        = strMsg & vbCrLf & "  SPN Details:"
      objAttr       = objAccount.Get("servicePrincipalName")
      For Each strAttr In objAttr
        strMsg      = strMsg & vbCRLF & "    " & strAttr
      Next
  End Select

  Select Case True
    Case IsNull(objAccount.Get("msDS-AllowedToDelegateTo"))
      ' Nothing, no SPN Usage for Account
    Case Else
      strMsg        = strMsg & vbCrLf & "  Delegation Details:"
      objAttr       = objAccount.Get("msDS-AllowedToDelegateTo")
      For Each strAttr In objAttr
        strMsg      = strMsg & vbCRLF & "    " & strAttr
      Next
  End Select

  If strMsg <> Mid(objAccount.Name, 4) Then
    Wscript.Echo " "
    Wscript.Echo strMsg
  End If

End Sub


Sub Terminate()

  objRecordset.Close
  objConnection.Close

  wscript.echo vbCrLf & "-- End of Report --"

  wscript.quit 0

End Sub
 