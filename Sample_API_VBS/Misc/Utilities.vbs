Option Explicit

Dim objNMUtils
Dim strTreeConfig

Set objNMUtils = CreateObject( "ActiveXperts.NMUtilities" )

WScript.Echo objNMUtils.GetNMRootDir
WScript.Echo objNMUtils.GetDateFormat
WScript.Echo objNMUtils.GetTimeFormat

WScript.Echo objNMUtils.RpGetToken( "2923BE8A8632929923A193239B973129" )
WScript.Echo objNMUtils.RpGetTraceFile 


WScript.Echo objNMUtils.RpGetFromSecs
WScript.Echo objNMUtils.RpGetToSecs

Const PERIOD_UNKNOWN	= 0 
Const PERIOD_DAY		= 1 
Const PERIOD_WEEK		= 2
Const PERIOD_MONTH		= 3 
Const PERIOD_QUARTER	= 4
Const PERIOD_YEAR		= 5
objNMUtils.RpSetPeriod PERIOD_DAY, 1

WScript.Echo objNMUtils.RpGetFromSecs
WScript.Echo objNMUtils.RpGetToSecs


WScript.Echo "Ready."





