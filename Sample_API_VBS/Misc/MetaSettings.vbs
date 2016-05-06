Option Explicit

Dim objNMConfig, objNMMeta

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objNMMeta   = objNMConfig.LoadMetaSettings

PrintMetaSettings objNMMeta  

objNMConfig.Close 

WScript.Echo "Ready."


Sub PrintMetaSettings( objNMMeta )
  WScript.Echo vbCrLf & "MetaSettings:"
  WScript.Echo "  ID: " & objNMMeta.ID   
  WScript.Echo "  Version: " & objNMMeta.Version   
  WScript.Echo "  Engine0Root: " & objNMMeta.Engine0Root   
  WScript.Echo "  Reserved1: " & objNMMeta.Reserved1   
  WScript.Echo "  Reserved2: " & objNMMeta.Reserved2   
 
  WScript.Echo 
 
End Sub






