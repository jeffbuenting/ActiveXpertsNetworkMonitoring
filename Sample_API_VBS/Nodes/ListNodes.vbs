Option Explicit

Dim objNMConfig, c
Dim objNode

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )
Set c           = CreateObject( "ActiveXperts.NMConstants" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

' Find all Folders
Set objNode = objNMConfig.FindFirstNode( "Type = " & c.CHECKTYPE_FOLDER )
If( objNMConfig.LastError <> 0 ) Then
   WScript.Echo "FindFirstNode, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
Else
   WScript.Echo vbCrLf & "All Folders: "
   While( objNMConfig.LastError = 0 )
      PrintNode( objNode )
      Set objNode = objNMConfig.FindNextNode
   WEnd
End If

' Find all ICMP Rules
Set objNode = objNMConfig.FindFirstNode( "Type = " & c.CHECKTYPE_ICMP)
If( objNMConfig.LastError <> 0 ) Then
   WScript.Echo "FindFirstNode, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
Else
   WScript.Echo vbCrLf & "All ICMP Rules: "
   While( objNMConfig.LastError = 0 )
      PrintNode( objNode )
      Set objNode = objNMConfig.FindNextNode
   WEnd
End If


' Find all Rules other than ICMP and other than Folders:
Set objNode = objNMConfig.FindFirstNode( "Type <> " & c.CHECKTYPE_ICMP & " AND Type <> " & c.CHECKTYPE_FOLDER )
If( objNMConfig.LastError <> 0 ) Then
   WScript.Echo "FindFirstNode, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
Else
   WScript.Echo vbCrLf & "All rules except ICMP rules and Folders: "
   While( objNMConfig.LastError = 0 )
      PrintNode( objNode )
      Set objNode = objNMConfig.FindNextNode
   WEnd
End If

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."


Sub PrintNode( objNode )
' This function only prints a few properties of the objNode object
' For a full list of properties, see also 'LoadNode.vbs', or check 
' the [Nodes] table in the [Config.mdb] database

  WScript.Echo "  Node [" & objNode.ID & "] : " & objNode.DisplayName
End Sub












