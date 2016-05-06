' ///////////////////////////////////////////////////////////////////////////////////////
'
' Every Rule has a unique ID in the [Nodes] table of the [Config.mdb] database
' The ID's of new rules start at NODEID_USERBASE
' Folders are stored as Rules in the database. 
' 
' There's one special Folder in the database, identified by ID 1: the Root folder.
'
' There's one special Rule in the database, identified by ID 2: the rule that holds the
' defualt values for all new rules. 
'
' ///////////////////////////////////////////////////////////////////////////////////////

Set c = CreateObject( "ActiveXperts.NMConstants" )

WScript.Echo "NODEID_UNDEFINED: " & c.NODEID_UNDEFINED
WScript.Echo "NODEID_ROOT: " & c.NODEID_ROOT
WScript.Echo "NODEID_DEFAULTSETTINGS: " & c.NODEID_DEFAULTSETTINGS
WScript.Echo "NODEID_USERBASE: " & c.NODEID_USERBASE


