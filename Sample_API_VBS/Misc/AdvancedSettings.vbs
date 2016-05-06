Option Explicit

Dim objNMConfig
Dim strSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
If( objNMConfig.LastError <> 0 ) Then
   PrintError( objNMConfig )
   WScript.Quit
End If
WScript.Echo "Database opened."

' Load ADVANCED-UNCERTAIN
strSettings = objNMConfig.LoadAdvancedSettings( 1 )
If( objNMConfig.LastError <> 0 ) Then
   PrintError objNMConfig
   objNMConfig.Close
   WScript.Quit
End If
WScript.Echo vbCrLf & "ADVANCED-UNCERTAIN: [" & strSettings & "]"

' Load ADVANCED-HTTP
strSettings = objNMConfig.LoadAdvancedSettings( 2 )
If( objNMConfig.LastError <> 0 ) Then
   PrintError objNMConfig
   objNMConfig.Close
   WScript.Quit
End If
WScript.Echo vbCrLf & "ADVANCED-HTTP: [" & strSettings & "]"

' Load ADVANCED-THREADS
strSettings = objNMConfig.LoadAdvancedSettings( 4 )
If( objNMConfig.LastError <> 0 ) Then
   PrintError objNMConfig
   objNMConfig.Close
   WScript.Quit
End If
WScript.Echo vbCrLf & "ADVANCED-THREADS: [" & strSettings & "]"

' Load ADVANCED-WEB VIEWS
strSettings = objNMConfig.LoadAdvancedSettings( 5 )
If( objNMConfig.LastError <> 0 ) Then
   PrintError objNMConfig
   objNMConfig.Close
   WScript.Quit
End If
WScript.Echo vbCrLf & "ADVANCED-WEB VIEWS: [" & strSettings & "]"


' Load ADVANCED-REPORT GENERATOR
strSettings = objNMConfig.LoadAdvancedSettings( 6 )
If( objNMConfig.LastError <> 0 ) Then
   PrintError objNMConfig
   objNMConfig.Close
   WScript.Quit
End If
WScript.Echo vbCrLf & "ADVANCED-REPORT GENERATOR: [" & strSettings & "]"


objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintError( objNMConfig )
   WScript.Echo "Operation failed, error: " & objNMConfig.LastError & _
  " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
End Sub






