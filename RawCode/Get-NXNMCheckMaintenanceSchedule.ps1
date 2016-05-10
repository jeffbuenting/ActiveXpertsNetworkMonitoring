Function Convert-DateFromUnix {

<#
    .SYnopsis
        Converts Unix DateTime to Powershell DateTime

    .Description
        Converts Unix (Epoch) date time to powershell Date Time.

    .Parameter DateTime
        Date/Time in Epoch format

    .Parameter Date
        Date in Epoch Format

    .Parameter Time
        Time in Epoch Format

    .links
        http://www.epochconverter.com/

    .Links
        https://nzfoo.wordpress.com/2014/01/21/converting-from-unix-timestamp/

    .Notes
        Author: Jeff Buenting
        Date: 2016 FEB 8
#>

    [CmdletBinding()]
    Param (
        [Parameter(ParameterSetName='DateTime',Mandatory=$True,ValueFromPipeline=$True)]
        [String[]]$DateTime,
        
        [Parameter(ParameterSetName='Date')] 
        [String]$Date,

        [Parameter(ParameterSetName='Date')] 
        [String]$Time
    )

    Process {
        Switch ( $PSCmdlet.ParameterSetName ) {
            'DateTime' {
                Write-Verbose 'DateTime Parameter Set'

                Foreach ( $D in $DateTime ) {
                    Write-Verbose "Converting $D"
                    write-Output ([timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($D)))
                }
            }

            'Date' {
                Write-Verbose 'Date Parameter Set'
               
                # ------ Return todays data if no date parameter was specified.
                 Write-Verbose "Converting $Date"
                if ( $Date ) {
                        $D = "{0:D}" -f ( [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Date)) )
                    }
                    else {
                        $D = "{0:D}" -f ( Get-Date )
                }

                Write-verbose "Date: $D"

                # ----- Return the time now if no time was specified
                Write-Verbose "Converting $Time"
                if ( $Time ) {
                        $T = "{0:T}" -f ( [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($Time)) )
                    }
                    else {
                        $T = "{0:T}" -f ( Get-Date )
                }

                Write-Verbose "Time: $T"

               

              

               Write-Output $(Get-Date "$D $T")
               
            }
        }
    }
}

#----------------------------------------------------------------------------------------


Function Convert-AXNMMaintScheduletoDate {

<#
    .Synopsis
        Convert Maintenace schedule to real date

    .Description
        Convert Maintenace schedule to real date.  Untility function  Not really needed to be exposed.

    .Parameter MaintDate
        Maintenance date in ActiveXperts NM format

    .Note
        Author: Jeff Buenting
        Date: 2016 FEB 9
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [String]$MaintDate
    )

    Process {
        Write-Verbose "Converting $MaintDate"

        $MD = $MaintDate -Split ';'

        Switch ( $MD[0].substring(0,1) ) {
            'e' {
                Write-Verbose "Every Day of the week"
                Switch ( $MD[0].substring(1) ) {
                    '1000000' {
                        Write-Verbose "     on Sunday"
                        $Day = 'Every Sunday'
                    }
                    '0100000' {
                        Write-Verbose "     on Monday"
                        $Day = 'Every Monday'
                    }
                    '0010000' {
                        Write-Verbose "     on Tuesday"
                        $Day = 'Every Tuesday'
                    }
                    '0001000' {
                        Write-Verbose "     on Wednesday"
                        $Day = 'Every Wednesday'
                    }
                    '0000100' {
                        Write-Verbose "     on Thursday"
                        $Day = 'Every Thursday'
                    }
                    '0000010' {
                        Write-Verbose "     on Friday"
                        $Day = 'Every Friday'
                    }
                    '0000001' {
                        Write-Verbose "     on Saturday"
                        $Day = 'Every Saturday'
                    }
                }
                $Day = "$Day @ {0:T}" -f ( Convert-DateFromUnix -Date $MD[0].substring(1) -Time $MD[1] ) 
            }

            'd' {
                Write-Verbose "Specific Date"
                $Day = '{0:F}' -f (Convert-DateFromUnix -Date $MD[0].substring(1) -Time $MD[1] )
            }
        }
    
        $MaintSched = New-object -TypeName psobject -Property @{    
            Date = $Day
            Duration = $MD[2]
        }
        #Write-Verbose $($MaintSched | out-string )
        Write-Output $MaintSched
    }
}

#----------------------------------------------------------------------------------------

Function Get-AXNMCheckMaintenanceSchedule {

<#
    .Synopsis
        Returns maintnance schedule for a Rule

    .Description
        Returns a ActiveXpert Network Monitoring Maintenance Schedule for a Rule.  Returns both the Global and Rule specific Schedules.

    .Parameter Rule
        Rule for which to get maintenance Schdules.  Use Get-AXNMRule.

    .Example
        get maintenance schedue for the ICMP rule

        Get-AXNMCheck  -Verbose | where Displayname -eq 'Server - ICMP Ping' | Get-AXNMCheckMaintenanceSchedule

    .Notes
        Author: Jeff Buenting
        Date: 2016 FEB 9
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [PSObject[]]$Rule
    )

    Begin {
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
        Foreach ( $NMC in $Rule ) {
            Write-Verbose "Getting Maintenance Schedule for $($NMC.DisplayName)"

            # -----  Get the Maintenance for the Overall tool
            ($NMConfig.LoadMaintenanceSettings()) -split '\|' | foreach {
                $MaintSched = $_ | Convert-AXNMMaintScheduletoDate
                $MaintSched | Add-Member -MemberType NoteProperty -Name Scope -Value Global
                Write-Output $MaintSched
            }

            # ----- Get the Schedule for the NM Check
            ($NMCheck.MaintenanceList) -Split '\|' | foreach {
                $MaintSched = $_ | Convert-AXNMMaintScheduletoDate
                $MaintSched | Add-Member -MemberType NoteProperty -Name Scope -Value Rule
                Write-Output $MaintSched
            }
        }
    }

    End {
        Write-Verbose "Closing the Network Monitoring Database"
        $NMConfig.Close()
    }
        
}

Get-AXNMCheck  -Verbose | where Displayname -eq 'Rwva-ts1 - ICMP Ping' | Get-AXNMCheckMaintenanceSchedule -Verbose