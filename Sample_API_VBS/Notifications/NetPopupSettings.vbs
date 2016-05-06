Option Explicit

Dim objNMConfig
Dim bNetPopupEnabled

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

bNetPopupEnabled = objNMConfig.LoadNetPopupSettings()
WScript.Echo "netpopup enabled: " & bNetPopupEnabled

PrintDistributionGroups( objNMConfig )

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintDistributionGroups( objNMConfig )
  Dim objDistrGroup
  Set objDistrGroup = objNMConfig.FindFirstDistributionGroup( "NotificationTypeID = 2" )
  While( objNMConfig.LastError = 0 )
      WScript.Echo vbCrLf & "Distribution Group:"
      WScript.Echo "  ID   : " & objDistrGroup.ID
      WScript.Echo "  Name : " & objDistrGroup.Name
      WScript.Echo "  Recipients : " & objDistrGroup.Recipients
      Set objDistrGroup = objNMConfig.FindNextDistributionGroup
  WEnd
  WScript.Echo 
End Sub






