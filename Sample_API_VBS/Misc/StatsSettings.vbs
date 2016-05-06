Option Explicit

Dim objNMConfig
Dim objStatsSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objStatsSettings = objNMConfig.LoadStatsSettings
WScript.Echo "LoadStatsSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintStatsSettings( objStatsSettings )

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintStatsSettings( objStatsSettings )
  WScript.Echo vbCrLf & "StatsSettings:"
  WScript.Echo "  ID: " & objStatsSettings.ID
  WScript.Echo "  Enabled: " & objStatsSettings.Enabled
  WScript.Echo "  ConnectionString: " & objStatsSettings.ConnectionString
  WScript.Echo "  ConnectionString_AsIs: " & objStatsSettings.ConnectionString_AsIs
  WScript.Echo "  ConnectionEPassword: " & objStatsSettings.ConnectionEPassword
  WScript.Echo "  ReportsLocalPath: " & objStatsSettings.ReportsLocalPath
  WScript.Echo "  ReportsLocalPath_AsIs: " & objStatsSettings.ReportsLocalPath_AsIs
  WScript.Echo "  ReportsPublicPath: " & objStatsSettings.ReportsPublicPath
  WScript.Echo "  ReportsPublicPath_AsIs: " & objStatsSettings.ReportsPublicPath_AsIs
  WScript.Echo "  GraphsLocalPath: " & objStatsSettings.GraphsLocalPath
  WScript.Echo "  GraphsLocalPath_AsIs: " & objStatsSettings.GraphsLocalPath_AsIs
  WScript.Echo "  DoCleanMonths: " & objStatsSettings.DoCleanMonths
  WScript.Echo "  CleanNumMonths: " & objStatsSettings.CleanNumMonths
  WScript.Echo 
End Sub






