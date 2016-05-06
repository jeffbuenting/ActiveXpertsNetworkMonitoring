Option Explicit

Dim objNMConfig
Dim objWebView
Dim i

Set objNMConfig = CreateObject( "ActiveXperts.NMConfig" )

objNMConfig.Open
WScript.Echo "Open, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
If( objNMConfig.LastError <> 0 ) Then
  WScript.Quit
End If

i = 0
While( objNMConfig.LastError = 0 )

   Set objWebView = objNMConfig.LoadWebView( i )
   WScript.Echo "LoadWebView, result: " & objNMConfig.LastError & " (" &  objNMConfig.GetErrorDescription( objNMConfig.LastError ) & ")"
   If( objNMConfig.LastError = 0 ) Then
      PrintWebView( objWebView )
   End If
   i = i + 1
WEnd

objNMConfig.Close
WScript.Echo "Database closed."

WScript.Echo "Ready."

Sub PrintWebView( objWebView )
  WScript.Echo vbCrLf & "WebView:"
  WScript.Echo "  ID: "       & objWebView.ID
  WScript.Echo "  Enabled: "  & objWebView.Enabled
  WScript.Echo "  Name: "     & objWebView.Name
  WScript.Echo "  XmlFile: "  & objWebView.XmlFile
  WScript.Echo "  XmlFile_AsIs: "  & objWebView.XmlFile_AsIs
  WScript.Echo "  XslFile: "  & objWebView.XslFile
  WScript.Echo "  XslFile_AsIs: "  & objWebView.XslFile_AsIs
  WScript.Echo "  Comments: " & objWebView.Comments
  WScript.Echo 
End Sub




