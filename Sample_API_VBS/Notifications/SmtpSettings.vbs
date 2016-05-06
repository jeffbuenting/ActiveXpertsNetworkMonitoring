Option Explicit

Dim objNMConfig
Dim objSmtpSettings

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objSmtpSettings = objNMConfig.LoadSmtpSettings( True )  ' True means: primary server 
WScript.Echo "LoadSmtpSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintSmtpSettings( objSmtpSettings )

Set objSmtpSettings = objNMConfig.LoadSmtpSettings( False)  ' False means: secundairy (fallback) SMTP server 
WScript.Echo "LoadSmtpSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  objNMConfig.Close
  WScript.Quit
End If

PrintSmtpSettings( objSmtpSettings )

PrintDistributionGroups( objNMConfig )

' Modify the SMTP settings
' objSmtpSettings.Server = "myserver@mydomain.com"
' objNMConfig.SaveSmtpSettings objSmtpSettings
' WScript.Echo "SaveSmtpSettings, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
' If( objNMConfig.LastError <> 0 ) Then
'   objNMConfig.Close
'  WScript.Quit
' End If

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintSmtpSettings( objSmtpSettings )
  WScript.Echo vbCrLf & "SmtpSettings:"
  WScript.Echo "  ID: " & objSmtpSettings.ID
  WScript.Echo "  Enabled: " & objSmtpSettings.Enabled
  WScript.Echo "  Host: " & objSmtpSettings.Host
  WScript.Echo "  Port: " & objSmtpSettings.Port
  WScript.Echo "  HasLogin: " & objSmtpSettings.HasLogin
  WScript.Echo "  Login: " & objSmtpSettings.Login
  WScript.Echo "  Password (encrypted): " & objSmtpSettings.EPassword
  WScript.Echo "  SpaAuthentication: " & objSmtpSettings.SpaAuthentication
  WScript.Echo "  SenderDisplayName: " & objSmtpSettings.SenderDisplayName
  WScript.Echo "  SenderEmail: " & objSmtpSettings.SenderEmail
  WScript.Echo 
End Sub


Sub PrintDistributionGroups( objNMConfig )
  Dim objDistrGroup
  Set objDistrGroup = objNMConfig.FindFirstDistributionGroup( "NotificationTypeID = 1" )
  While( objNMConfig.LastError = 0 )
      WScript.Echo vbCrLf & "Distribution Group:"
      WScript.Echo "  ID   : " & objDistrGroup.ID
      WScript.Echo "  Name : " & objDistrGroup.Name
      WScript.Echo "  Recipients : " & objDistrGroup.Recipients
      Set objDistrGroup = objNMConfig.FindNextDistributionGroup
  WEnd
  WScript.Echo 
End Sub






