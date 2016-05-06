Option Explicit

Dim objNMConfig
Dim strTreeConfig

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

WScript.Echo "objNMConfig.ConfigDatabase:" & objNMConfig.ConfigDatabase
WScript.Echo "objNMConfig.ReportDatabase:" & objNMConfig.ReportDatabase

strTreeConfig = objNMConfig.LoadTreeConfiguration
If( objNMConfig.LastError <> 0 ) Then
  WScript.Echo "LoadTreeConfiguration, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
  objNMConfig.Close
  WScript.Quit
End If

WScript.Echo "objNMConfig.LoadTreeConfiguration(): " & strTreeConfig

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."





