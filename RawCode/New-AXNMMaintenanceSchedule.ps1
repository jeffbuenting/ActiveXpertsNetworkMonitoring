Function New-AXNMMaintenanceSchedule {

<# 
    .Synopsis
        Creates a New Maintenance Schedule.

    .Description
        Creates a new Maintenance Schedule ( Global or Rule Specific ) in ActiveXperts Network Monitor.

    .Parameter Rule
        ActiveXpert Network Monitor Rule

    .Parameter MaintenanceRule
        Maintenance Schedule in a readable format.

        Every Day @ 1:36 PM
        Thursday, February 8, 2016 10:22:18 AM
        Wednesday, October 10, 2025 11:34 PM

    .Parameter Duration
        How many hours the schedule is for.

    .Example
        create a maintenance Schedule for one rule.

        Get-AXNMCheck  -Verbose | where Displayname -eq 'Server - ICMP Ping' | New-AXNMMaintenanceSchedule -MaintenanceSched 'Every Thursday @ 12:00:00 AM' -Duration 2 

    .Example
        Create a Global Maintenance Schedule

        New-AXNMMaintenanceSchedule -date 'Every Thursday @ 12:00:00 AM' -Duration 2 

    .Link
        Author: Jeff Buenting
        Date: 2016 FEB 10
#>

    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$true)]
        [PSObject]$Rule,

        [Parameter(Mandatory=$True)]
        [ValidateScript( {
            $D = 0
            #($_.tolower() -in 'every sunday','every monday','every tuesday','every wednesday','every thursday','every friday','every saturday') -or ([datetime]::Tryparse($_,[ref]$D))
            ($_.tolower() -match 'every .*day @ \d+:\d+(:\d+)? [a|p]m') -or ([datetime]::Tryparse($_,[ref]$D))
        } ) ]
        [String]$MaintenanceSched,

        [Parameter(Mandatory=$True)]
        [String]$Duration
    )

    Begin {
        # ----- Convert Date and Time to values ActiveXpert Understands (unix time)
        $DT = $Date -split '@'
        
        Switch ( $DT[0].tolower() ) {
            'every sunday ' {
                $Day = 'e1000000'
                $T = $DT[1]
            }
            'every monday ' {
                $Day = 'e0100000'
                $T = $DT[1]
            }
            'every tuesday ' {
                $Day = 'e0010000'
                $T = $DT[1]
            }
            'every wednesday ' {
                $Day = 'e0001000'
                $T = $DT[1]
            }
            'every thursday ' {
                $Day = 'e0000100'
                $T = $DT[1]
                      }
            'every friday ' {
                $Day = 'e0000010'
                $T = $DT[1]
            }
            'every saturday ' {
                $Day = 'e000001'
                $T = $DT[1]
            }
            default {
                $Day = New-TimeSpan -Start "01/01/1970 00:00" -End (Get-Date $Date).Date | Select-Object -ExpandProperty TotalSeconds
                $T = ( "{0:T}" -f (Get-Date $Date ) )
            }
        }

        # ----- ActiveXpert Net work Monitoring time is measured in Epoch time.  However for some reason, midnight is Jan 3, 1970 midnight.  UTC.  So convert and mesure from those dates to get the time.
        $Time = New-TimeSpan -Start "01/01/1970 00:00" -End ((Get-Date "Saturday, January 3, 1970$T").ToUniversalTime()) | Select-Object -ExpandProperty TotalSeconds

        $NewMaintSched = "$Day;$Time;$Duration"

        Write-Verbose "New Maintenance Schedule: $NewMaintSched"
   
        # ----- Get the NM Config object and open the database
        Write-Verbose "Open the Network MOnitoring Database"
        Try {
                $NMConfig = New-Object -ComObject ActiveXperts.NMConfig -ErrorAction Stop 
            }
            catch {
                Throw "Get-AXNMRule : $($_.Exception.message)`nCheck if ActiveXperts Network Monitoring is installed on $env:ComputerName"
        }
        $NMConfig.Open()
    }

    Process {
        if ( $Rule ) {
                Write-Verbose "Adding Maintenace schedule for rule: $($Rule.Displayname)"
                # ----- Load Rule from the Database modify and save
                $Node = $NMConfig.LoadNode( $Rule.ID )
                $Node.Maintenancelist += "|$NewMaintSched"
                $NMConfig.SaveNode( $Node )
                
            }
            else {
                Write-Verbose "Adding Global Maintenance Schedule"
                # ----- Load schedules from the Database modify and save
                $Sched = $NMConfig.LoadMaintenanceSettings()
                $Sched += "|$NewMaintSched"
                $NMConfig.SaveMaintenanceSettings( $Sched )
        }
    }

    End {
        Write-Verbose "Closing the Network Monitoring Database"
        $NMConfig.Close()
    }

}

#Get-AXNMCheck  -Verbose | where Displayname -eq 'Rwva-ts1 - ICMP Ping' | New-AXNMMaintenanceSchedule -date 'Monday, February 8, 2016 12:00:00 AM' -Duration 2 -Verbose
Get-AXNMCheck  -Verbose | where Displayname -eq 'Rwva-ts1 - ICMP Ping' | New-AXNMMaintenanceSchedule -date 'Every Thursday @ 12:00:00 AM' -Duration 2 -Verbose
#-date 'Every Sunday' -verbose

#-date 'Monday, February 8, 2016 12:00:00 AM' -Verbose