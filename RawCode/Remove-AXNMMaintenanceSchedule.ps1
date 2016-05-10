Function Remove-AXNMMaintenanceSchedule {

<#
    .Synopsis
        Removes a maintenance schedule from ActiveXperts Network Monitor

    .Description
        Removes a maintenance schedule from ActiveXperts Network Monitor.  Can remove either gobla or rule specific maintenance rules

    .Parameter MaintenanceSchedule
        Maintenance Schedule to remove.  Use Get-AXNMMaintenanceSchedule to obtain the mainenance object to remove.

    .Parameter Rule
        Rule to remove the maintenance schedule from.  If blank the rules will be removed from the global list.

    .Example
        Removes all Maintenance schedules from the Server - ICMP Ping rule dated February 10

        $Rule = Get-AXNMRule | where Displayname -eq 'Rwva-ts1 - ICMP Ping' 
        $MS = $Rule | Get-AXNMMaintenanceSchedule  | where Date -like "*February 10*"  
        $MS | Remove-AXNMMaintenanceSchedule -Rule $Rule -Verbose
        
    .Notes
        Author: Jeff Buenting
        Date: 2016 FEB 16
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [PSObject[]]$MaintenanceSchedule,

        [PSobject]$Rule
    )

   Begin {
        # ----- Get the NM Config object and open the database
        Write-Verbose "Open the Network MOnitoring Database"
        $NMConfig = New-Object -ComObject ActiveXperts.NMConfig
        
        $NMConfig.Open()
    } 

    Process {
        foreach ( $M in $MaintenanceSchedule ) {
            Write-Verbose "Removing Maintenance Schedule $($M.Date)"
           # $MS = Convert-AXNMDatetoMaintenanceSchedule -MaintenanceSched $M.Date -Duration $M.Duration
           # Write-Verbose "     Converted $MS"

            if ( $Rule ) {
                    Write-Verbose "Removing Maintenace schedule for rule: $($Rule.Displayname)"

                    # ----- Check the Maintenance Schedules scope
                    if ( $M.Scope -eq $Rule.Displayname ) {
                            Write-Verbose "Valid Schedule for Rule, Removing"

                            # ----- Load Rule from the Database modify and save
                            $Node = $NMConfig.LoadNode( $Rule.ID )

                            # ----- Split the schedule and Loop thru and skip the schedule being deleted
                            $List = $NUll
                            ($Node.Maintenancelist) -split '\|' | Foreach {
                                $MS = $_ | Convert-AXNMMaintScheduletoDate
                                 
                                if ( ($MS.Date -ne $M.date) -and ($MS.Duration -ne $M.Duration ) ) {
                                    Write-Verbose "     Keeping $MS"
                                    # ----- Add divider if list is not null
                                    if ( $List -ne $Null ) { $List += '|' }
                                    $List += "$_"
                                }
                            }
                           
                            $Node.MaintenanceList = $List
                                          
                            $NMConfig.SaveNode( $Node )
                        }
                        Else {
                            Throw "Remove-AXNMMaintenanceSchedule : Cannot remove a Global schedule from a Rule. `n$_.Exception.Message"
                    }
                
                }
                else {
                    Write-Verbose "Removing Global Maintenance Schedule"

                    # ----- Check the Maintenance Schedules Scope
                    if ( $M.Scope -eq "Global" ) {
                            Write-Verbose "Valid Scope, Removing"

                            # ----- Load schedules from the Database modify and save
                            # ----- Split the schedule and Loop thru and skip the schedule being deleted
                            $List = $NUll
                            $Sched = $NMConfig.LoadMaintenanceSettings()
                            ($Sched.LoadMaintenanceSettings()) -split '\|' | Foreach {
                                $MS = $_ | Convert-AXNMMaintScheduletoDate
                                 
                                if ( ($MS.Date -ne $M.date) -and ($MS.Duration -ne $M.Duration ) ) {
                                    Write-Verbose "     Keeping $MS"
                                    # ----- Add divider if list is not null
                                    if ( $List -ne $Null ) { $List += '|' }
                                    $List += "$_"
                                }
                            }
                         
                            $Sched = $List
                            $NMConfig.SaveMaintenanceSettings( $Sched )
                        }
                        else {
                            Throw "Remove-AXNMMaintenanceSchedule : Cannot remove a Rule schedule Globally.`n$($_.Exception.Message)"
                    }
            }
            Write-Verbose "-----"
        }
    }

    End {
        Write-Verbose "Closing the Network Monitoring Database"
        $NMConfig.Close()
    }

}

$Rule = Get-AXNMRule | where Displayname -eq 'Rwva-ts1 - ICMP Ping' 
$MS = $Rule | Get-AXNMMaintenanceSchedule  | where Date -like "*February 10*" 
$MS
$MS | foreach {
     Convert-AXNMDatetoMaintenanceSchedule -MaintenanceSched $_.Date -Duration $_.Duration -verbose
}
$MS | Remove-AXNMMaintenanceSchedule -Rule $Rule -Verbose