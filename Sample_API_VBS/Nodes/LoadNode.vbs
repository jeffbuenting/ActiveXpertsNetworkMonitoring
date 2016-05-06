Option Explicit

Dim objNMConfig, c
Dim objNode

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )
Set c           = CreateObject( "ActiveXperts.NMConstants" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

' Load Check with ID=10000
Set objNode = objNMConfig.LoadNode( 10000 )
If( objNMConfig.LastError = 0 ) Then
   PrintNode( objNode )
Else
   WScript.Echo "LoadNode ERROR: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
End If

' Load a special Folder: the Root folder
Set objNode = objNMConfig.LoadNode( c.NODEID_ROOT )
If( objNMConfig.LastError = 0 ) Then
   PrintNode( objNode )
Else
   WScript.Echo "LoadNode ERROR: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
End If


objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."




Sub PrintNode( objNode )
  WScript.Echo vbCrLf & "Node [" & objNode.ID & "] : "
  WScript.Echo "  ID: " & objNode.ID
  WScript.Echo "  DisplayName: " & objNode.DisplayName
  WScript.Echo "  Comments: " & objNode.Comments
  WScript.Echo "  ScanInterval: " & objNode.ScanInterval
  WScript.Echo "  ErrorTreshold: " & objNode.ErrorTreshold
  WScript.Echo "  OnHold: " & objNode.OnHold 
  WScript.Echo "  TimeOut: " & objNode.TimeOut
  WScript.Echo "  Type: " & objNode.Type

  WScript.Echo "  MaintenanceServer: " & objNode.MaintenanceServer
  WScript.Echo "  MaintenanceList: " & objNode.MaintenanceList
  WScript.Echo "  DependencyList: " & objNode.DependencyList
  WScript.Echo "  NotifyMultiple: " & objNode.NotifyMultiple
  WScript.Echo "  NotifyMinutes: " & objNode.NotifyMinutes
  WScript.Echo "  NotifyFlags: " & objNode.NotifyFlags
  WScript.Echo "  NotifyMailOffline: " & objNode.NotifyMailOffline
  WScript.Echo "  NotifyMailOnline: " & objNode.NotifyMailOnline
  WScript.Echo "  NotifyNetworkOffline: " & objNode.NotifyNetworkOffline
  WScript.Echo "  NotifyNetworkOnline: " & objNode.NotifyNetworkOnline
  WScript.Echo "  NotifySmsOffline: " & objNode.NotifySmsOffline
  WScript.Echo "  NotifySmsOnline: " & objNode.NotifySmsOnline
  WScript.Echo "  NotifyPagerOffline: " & objNode.NotifyPagerOffline
  WScript.Echo "  NotifyPagerOnline: " & objNode.NotifyPagerOnline
  WScript.Echo "  NotifySnmpTrapOffline: " & objNode.NotifySnmpTrapOffline
  WScript.Echo "  NotifySnmpTrapOnline: " & objNode.NotifySnmpTrapOnline
  WScript.Echo "  RunMultiple: " & objNode.RunMultiple
  WScript.Echo "  RunMinutes: " & objNode.RunMinutes
  WScript.Echo "  RunFlags: " & objNode.RunFlags
  WScript.Echo "  RunExeOffline: " & objNode.RunExeOffline
  WScript.Echo "  RunExeOffline_AsIs: " & objNode.RunExeOffline_AsIs
  WScript.Echo "  RunExeOnline: " & objNode.RunExeOnline
  WScript.Echo "  RunExeOnline_AsIs: " & objNode.RunExeOnline_AsIs
  WScript.Echo "  RunVbsOffline: " & objNode.RunVbsOffline
  WScript.Echo "  RunVbsOffline_AsIs: " & objNode.RunVbsOffline_AsIs
  WScript.Echo "  RunVbsOnline: " & objNode.RunVbsOnline
  WScript.Echo "  RunVbsOnline_AsIs: " & objNode.RunVbsOnline_AsIs
  WScript.Echo "  RestartService: " & objNode.RestartService
  WScript.Echo "  RestartServer: " & objNode.RestartServer
  WScript.Echo "  RestartServerEntry: " & objNode.RestartServerEntry
  WScript.Echo "  sysParentID: " & objNode.sysParentID
  WScript.Echo "  CheckHasLogin1: " & objNode.CheckHasLogin1
  WScript.Echo "  CheckLoginName1: " & objNode.CheckLoginName1
  WScript.Echo "  CheckLoginEPassword1: " & objNode.CheckLoginEPassword1
  WScript.Echo "  CheckHasLogin2: " & objNode.CheckHasLogin2
  WScript.Echo "  CheckLoginName2: " & objNode.CheckLoginName2
  WScript.Echo "  CheckLoginEPassword2: " & objNode.CheckLoginEPassword2
  WScript.Echo "  CheckFlags: " & objNode.CheckFlags
  WScript.Echo "  CheckServer: " & objNode.CheckServer
  WScript.Echo "  CheckParam1: " & objNode.CheckParam1
  WScript.Echo "  CheckParam1_AsIs: " & objNode.CheckParam1_AsIs
  WScript.Echo "  CheckParam2: " & objNode.CheckParam2
  WScript.Echo "  CheckParam2_AsIs: " & objNode.CheckParam2_AsIs
  WScript.Echo "  CheckParam3: " & objNode.CheckParam3
  WScript.Echo "  CheckParam3_AsIs: " & objNode.CheckParam3_AsIs
  WScript.Echo "  CheckParam4: " & objNode.CheckParam4
  WScript.Echo "  CheckParam4_AsIs: " & objNode.CheckParam4_AsIs
  WScript.Echo "  CheckParam5: " & objNode.CheckParam5
  WScript.Echo "  CheckParam5_AsIs: " & objNode.CheckParam5_AsIs
  WScript.Echo "  CheckParam6: " & objNode.CheckParam6
  WScript.Echo "  CheckParam6_AsIs: " & objNode.CheckParam6_AsIs
  WScript.Echo "  CheckParam7: " & objNode.CheckParam7
  WScript.Echo "  CheckParam7_AsIs: " & objNode.CheckParam7_AsIs
  WScript.Echo "  CheckParam8: " & objNode.CheckParam8
  WScript.Echo "  CheckParam8_AsIs: " & objNode.CheckParam8_AsIs



  WScript.Echo "  sysParentID: " & objNode.sysParentID
  WScript.Echo 
End Sub







