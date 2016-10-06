# ActiveXperts_Network_Monitor_PowershellModule
Powershell Module for ActiveXperts Network Monitor Manager

This module can be used to interact with ActiveXperts Network Monitor Manager.  The original idea behind this was to develope an easy way to automate putting a monitor rule into maintenance mode.

Module contains the following Cmdlets

Get-AXNMRule
Get-AXNMMaintenanceSchedule
New-AXNMMaintenanceSchedule
Remove-AXNMAMaintenanceSchedule

For the life of me I cannot figure out why adding a maintenance schedule to an individual rule does not work.  It looks the same as when I do it via the gui.  Will have to work on that some day.

But, the New-AXNMMaintenanceschedule does set global manitance.  and Removes it.  SO at least I got that going for me.

To run remotely wrap the commands in Invoke-Command


```Powershell
$MaintSched = invoke-command -Computername ServerA -scriptblock { 
    import-module ActiveXpertsNetworkMonitoring -force
    $S = New-AXNMMaintenanceSchedule -MaintenanceSched (Get-Date) -Duration 3 -Passthru 
    Write-Output $S
}

$MaintSched

invoke-command -Computername ServerA -scriptblock {
    import-module ActiveXpertsNetworkMonitoring -force
    $Using:MaintSched | Remove-AXNMMaintenanceSchedule 
}
```

where ServerA is the server that ActiveXpert Network Monitor is installed.
