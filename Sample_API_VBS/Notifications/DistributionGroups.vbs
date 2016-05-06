Option Explicit

Const NOTIFICATIONTYPE_SMTP     = 1
Const NOTIFICATIONTYPE_NETPOPUP = 2
Const NOTIFICATIONTYPE_SMS      = 3
Const NOTIFICATIONTYPE_PAGER    = 4


Dim objNMConfig
Dim objDistrGroup

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

Set objDistrGroup = objNMConfig.LoadDistributionGroup( NOTIFICATIONTYPE_SMTP )
WScript.Echo "LoadDistributionGroup, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
PrintDistributionGroup "SMTP", objDistrGroup

Set objDistrGroup = objNMConfig.LoadDistributionGroup( NOTIFICATIONTYPE_NETPOPUP )
WScript.Echo "LoadDistributionGroup, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
PrintDistributionGroup "NETPOPUPP", objDistrGroup

Set objDistrGroup = objNMConfig.LoadDistributionGroup( NOTIFICATIONTYPE_SMS )
WScript.Echo "LoadDistributionGroup, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
PrintDistributionGroup "SMS", objDistrGroup

Set objDistrGroup = objNMConfig.LoadDistributionGroup( NOTIFICATIONTYPE_PAGER )
WScript.Echo "LoadDistributionGroup, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
PrintDistributionGroup "PAGER", objDistrGroup



objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."

Sub PrintDistributionGroup( strType, objDistrGroup )
  WScript.Echo vbCrLf & strType & " DistributionGroups:"
  WScript.Echo "  ID: "                 & objDistrGroup.ID
  WScript.Echo "  IsDefault: "          & objDistrGroup.IsDefault
  WScript.Echo "  DistributionTypeID: " & objDistrGroup.DistributionTypeID
  WScript.Echo "  Name: "               & objDistrGroup.Name
  WScript.Echo "  Recipients: "         & objDistrGroup.Recipients
  WScript.Echo 
End Sub




