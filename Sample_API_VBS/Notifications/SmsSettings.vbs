Option Explicit

Dim objNMConfig
Dim objSmsSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objSmsSettings = objNMConfig.LoadSmsSettings
WScript.Echo "LoadSmsSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintSmsSettings( objSmsSettings )

PrintDistributionGroups( objNMConfig )

' Modify the SMS settings
' objSmsSettings.SmscDevice = "Standard 9600 bps Modem"
' objNMConfig.SaveSmsSettings objSmsSettings
' WScript.Echo "SaveSmsSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
' If( objNMConfig.LastError <> 0 ) Then
'  objNMConfig.Close
'  WScript.Quit
' End If

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintSmsSettings( objSmsSettings )
  WScript.Echo vbCrLf & "SmsSettings:"
  WScript.Echo "  ID: " & objSmsSettings.ID
  WScript.Echo "  Enabled: " & objSmsSettings.Enabled
  WScript.Echo "  Protocol: " & objSmsSettings.Protocol
  WScript.Echo "  GsmHasPincode: " & objSmsSettings.GsmHasPincode
  WScript.Echo "  GsmPincode: " & objSmsSettings.GsmPincode
  WScript.Echo "  GsmDevice: " & objSmsSettings.GsmDevice
  WScript.Echo "  GsmDeviceInitString: " & objSmsSettings.GsmDeviceInitString
  WScript.Echo "  GsmDeviceSpeed: " & objSmsSettings.GsmDeviceSpeed
  WScript.Echo "  GsmDeviceFlowControl: " & objSmsSettings.GsmDeviceFlowControl
  WScript.Echo "  SmsServerChannel: " & objSmsSettings.SmsServerChannel
  WScript.Echo 
End Sub


Sub PrintDistributionGroups( objNMConfig )
  Dim objDistrGroup
  Set objDistrGroup = objNMConfig.FindFirstDistributionGroup( "NotificationTypeID = 3" )
  While( objNMConfig.LastError = 0 )
      WScript.Echo vbCrLf & "Distribution Group:"
      WScript.Echo "  ID   : " & objDistrGroup.ID
      WScript.Echo "  Name : " & objDistrGroup.Name
      WScript.Echo "  Recipients : " & objDistrGroup.Recipients
      Set objDistrGroup = objNMConfig.FindNextDistributionGroup
  WEnd
  WScript.Echo 
End Sub






