Option Explicit

Dim objNMConfig, numID

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

' Find all Folders
numID = objNMConfig.FindFirstNodeID( "" )
If( objNMConfig.LastError <> 0 ) Then
   WScript.Echo "FindFirstNodeID, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
Else
   WScript.Echo vbCrLf & "Rules: "
   While( objNMConfig.LastError = 0 )
      WScript.Echo "  ID: " & numID
      numID = objNMConfig.FindNextNodeID
   WEnd
End If

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."













