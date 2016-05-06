Option Explicit

Dim objNMConfig
Dim objPagerSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objPagerSettings = objNMConfig.LoadPagerSettings
WScript.Echo "LoadPagerSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintPagerSettings( objPagerSettings )

PrintDistributionGroups( objNMConfig )

' Modify the Pager settings
' objPagerSettings.Device = "Standard 1200 bps Modem"
objNMConfig.SavePagerSettings objPagerSettings
WScript.Echo "SavePagerSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintPagerSettings( objPagerSettings )
  WScript.Echo vbCrLf & "PagerSettings:"
  WScript.Echo "  ID: " & objPagerSettings.ID
  WScript.Echo "  Enabled: " & objPagerSettings.Enabled
  WScript.Echo "  Protocol: " & objPagerSettings.Protocol
  WScript.Echo "  DtmfDevice: " & objPagerSettings.DtmfDevice
  WScript.Echo "  DtmfDeviceSpeed: " & objPagerSettings.DtmfDeviceSpeed
  WScript.Echo "  DtmfDeviceFlowControl: " & objPagerSettings.DtmfDeviceFlowControl
  WScript.Echo "  DtmfDeviceInitString: " & objPagerSettings.DtmfDeviceInitString
  WScript.Echo "  DtmfDeviceTone: " & objPagerSettings.DtmfDeviceTone
  WScript.Echo "  DtmfDialPrefix: " & objPagerSettings.DtmfDialPrefix
  WScript.Echo "  DtmfRedialAttempts: " & objPagerSettings.DtmfRedialAttempts
  WScript.Echo "  SnppServer: " & objPagerSettings.SnppServer
  WScript.Echo "  SnppPort: " & objPagerSettings.SnppPort
  WScript.Echo "  SnppHasPassword: " & objPagerSettings.SnppHasPassword
  WScript.Echo "  SnppPassword (encrypted): " & objPagerSettings.SnppEPassword
  WScript.Echo "  SnppTimeOut: " & objPagerSettings.SnppTimeOut

  WScript.Echo 
End Sub


Sub PrintDistributionGroups( objNMConfig )
  Dim objDistrGroup
  Set objDistrGroup = objNMConfig.FindFirstDistributionGroup( "NotificationTypeID = 4" )
  While( objNMConfig.LastError = 0 )
      WScript.Echo vbCrLf & "Distribution Group:"
      WScript.Echo "  ID   : " & objDistrGroup.ID
      WScript.Echo "  Name : " & objDistrGroup.Name
      WScript.Echo "  Recipients : " & objDistrGroup.Recipients
      Set objDistrGroup = objNMConfig.FindNextDistributionGroup
  WEnd
  WScript.Echo 
End Sub






