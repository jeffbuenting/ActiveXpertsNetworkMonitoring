' ///////////////////////////////////////////////////////////////////////////////////////
'
' Every Event has a Result value in the [Events] table of the [ReportData.mdb] database
' The range of results are listed in the [Results] table of the [ReportData.mdb] database.
' To use such a constant, you can either use these database values, or use the constants
' listed below
'
' ///////////////////////////////////////////////////////////////////////////////////////

Set c = CreateObject( "ActiveXperts.NMConstants" )

WScript.Echo "RESULT_UNCERTAIN: " & c.RESULT_UNCERTAIN
WScript.Echo "RESULT_SUCCESS: " & c.RESULT_SUCCESS
WScript.Echo "RESULT_ERROR: " & c.RESULT_ERROR
WScript.Echo "RESULT_FAILURE: " & c.RESULT_FAILURE
WScript.Echo "RESULT_MAINTENANCE: " & c.RESULT_MAINTENANCE
WScript.Echo "RESULT_ONHOLD: " & c.RESULT_ONHOLD
WScript.Echo "RESULT_DEPENDEE_ERROR: " & c.RESULT_DEPENDEE_ERROR
WScript.Echo "RESULT_DEPENDEE_FAILURE: " & c.RESULT_DEPENDEE_FAILURE
WScript.Echo "RESULT_NOTPROCESSED: " & c.RESULT_NOTPROCESSED

