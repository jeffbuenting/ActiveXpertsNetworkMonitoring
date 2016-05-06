Option Explicit

Dim objNMConfig
Dim objLogSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objLogSettings = objNMConfig.LoadLogSettings
WScript.Echo "LoadLogSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintLogSettings( objLogSettings )

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintLogSettings( objLogSettings )
  WScript.Echo vbCrLf & "LogSettings:"
  WScript.Echo "  ID: " & objLogSettings.ID
  WScript.Echo "  EnableLog: " & objLogSettings.EnableLog
  WScript.Echo "  UseADO: " & objLogSettings.UseADO
  WScript.Echo "  LogSystemOnly: " & objLogSettings.LogSystemOnly
  WScript.Echo "  FileLogDir: " & objLogSettings.FileLogDir
  WScript.Echo "  FileLogDir_AsIs: " & objLogSettings.FileLogDir_AsIs  
  WScript.Echo "  FileLogSeparator: " & objLogSettings.FileLogSeparator
  WScript.Echo "  FileMaxSizeKB: " & objLogSettings.FileMaxSizeKB
  WScript.Echo "  FileBackupIfRequired: " & objLogSettings.FileBackupIfRequired
  WScript.Echo "  AdoDatabaseType: " & objLogSettings.AdoDatabaseType
  WScript.Echo "  AdoDatabaseConnection: " & objLogSettings.AdoDatabaseConnection
  WScript.Echo "  AdoDatabaseConnection_AsIs: " & objLogSettings.AdoDatabaseConnection_AsIs
  WScript.Echo "  AdoDatabaseEPassword: " & objLogSettings.AdoDatabaseEPassword
  WScript.Echo "  EnableSyslog: " & objLogSettings.EnableSyslog
  WScript.Echo "  SyslogMessage: " & objLogSettings.SyslogMessage
  WScript.Echo "  SyslogHost: " & objLogSettings.SyslogHost
  WScript.Echo "  SyslogPort: " & objLogSettings.SyslogPort
  WScript.Echo "  SyslogFacility: " & objLogSettings.SyslogFacility
  WScript.Echo "  SyslogPriorityFailure: " & objLogSettings.SyslogPriorityFailure
  WScript.Echo "  SyslogPriorityRecovery: " & objLogSettings.SyslogPriorityRecovery
  WScript.Echo "  SyslogPriorityInformation: " & objLogSettings.SyslogPriorityInformation
  WScript.Echo 
End Sub






