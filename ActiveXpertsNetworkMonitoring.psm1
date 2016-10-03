#----------------------------------------------------------------------------------
# AxtiveXperts Network Monitoring Powershell Module
#
# Author: Jeff Buenting
#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------
# Rule (nodes/checks) Cmdlets
#----------------------------------------------------------------------------------

Function Get-AXNMRule {

<#
    .Synopsis
        Returns ActiveXpert Network Monitoring Rule

    .Description
        Returns a monitoring rule from ActiveXpert Network Monitoring.  Rules are also called Checks or Nodes.

    .Parameter RuleType
        Used to filter the results by type of rule.

    .Examples
        Return all rules.

        Get-AXNMRule

    .Examples
        Return only the ICMP Rules

        Get-AXNMRule -RuleType CHECKTYPE_ICMP

    .Note
        Author: Jeff Buenting
        Date: 2016 FEB 11
#>

    [CmdletBinding()]
    Param (
        [Parameter (ParameterSetName = "CheckType")]
        [ValidateSet( 'ALL','CHECKTYPE_ADOSQL','CHECKTYPE_CPU','CHECKTYPE_DIRSIZE','CHECKTYPE_DISKS','CHECKTYPE_DISKSPACE','CHECKTYPE_DNS','CHECKTYPE_DOOR','CHECKTYPE_EVENTLOG','CHECKTYPE_FILE','CHECKTYPE_FLOPPY','CHECKTYPE_FOLDER','CHECKTYPE_FTP','CHECKTYPE_HTTP','CHECKTYPE_HUMIDITY','CHECKTYPE_ICMP','CHECKTYPE_IMAP','CHECKTYPE_LIGHT','CHECKTYPE_MEMORY','CHECKTYPE_MOTION','CHECKTYPE_MSMQ','CHECKTYPE_MSTSE','CHECKTYPE_NNTP','CHECKTYPE_NTP','CHECKTYPE_ODBC','CHECKTYPE_ORACLE','CHECKTYPE_POP3','CHECKTYPE_POWER','CHECKTYPE_PRINTER','CHECKTYPE_PROCESS','CHECKTYPE_REGISTRY','CHECKTYPE_RESISTANCE','CHECKTYPE_RSH','CHECKTYPE_SERVICE','CHECKTYPE_SMOKE','CHECKTYPE_SMTP','CHECKTYPE_SNMPGET','CHECKTYPE_SNMPTRAPRECEIVE','CHECKTYPE_SSH','CHECKTYPE_SWITCHNC','CHECKTYPE_SWITCHNO','CHECKTYPE_TCPIP','CHECKTYPE_TEMPERATURE','CHECKTYPE_UNDEFINED','CHECKTYPE_VBSCRIPT','CHECKTYPE_WETNESS' )]
        [String[]]$CheckType = 'ALL',

        [Parameter (ParameterSetName = "ID",Mandatory = $True)]
        [String]$ID

    )

    # ----- Get the NM Config object and open the database
    Write-Verbose "Open the Network MOnitoring Database"
    Try {
            $NMConfig = New-Object -ComObject ActiveXperts.NMConfig -ErrorAction Stop 
        }
        catch {
            Throw "Get-AXNMRule : $($_.Exception.message)`nCheck if ActiveXperts Network Monitoring is installed on $env:ComputerName"
    }
    $NMConstants = New-Object -ComObject ActiveXperts.NMConstants

    $NMConfig.Open()
    
    Switch ( $PSCmdlet.ParameterSetName ) {
        'CheckType' {
            foreach ( $Check in $CheckType ) {
                Write-Verbose "Getting the following Rule Types:"
                if ( $Check.Tolower() -eq 'all' ) {
                        Write-Verbose "----- All"
                        $NMCheck = $NMConfig.FindFirstNode( "Type >= 0" )
                    }
                    else {
                        Write-Verbose "----- $Check"
                    
                        $NMNode = $NMConfig.FindFirstNode( "Type = $($NMConstants."$Check")" )
                }

                Do {
                    Write-Output $NMNode
                    $NMNode = $NMConfig.FindNextNode()
                } While ( $NMConfig.LastError -eq 0 )
            }
        }

        'ID' {
            $NMNode = $NMConfig.FindFirstNode( "ID = $ID" )
            Write-Output $NMNode
        }
    }
   
    Write-Verbose "Closing the Network Monitoring Database"
    $NMConfig.Close()
}

#----------------------------------------------------------------------------------
# Maintenance Schedule Cmdlets
#----------------------------------------------------------------------------------

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

#----------------------------------------------------------------------------------

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

#----------------------------------------------------------------------------------

Function Get-AXNMMaintenanceSchedule {

<#
    .Synopsis
        Returns maintnance schedule for a Rule

    .Description
        Returns a ActiveXpert Network Monitoring Maintenance Schedule for a Rule.  Returns both the Global and Rule specific Schedules.

    .Parameter Rule
        Rule for which to get maintenance Schdules.  Use Get-AXNMRule.

    .Example
        get maintenance schedue for the ICMP rule

        Get-AXNMRule  -Verbose | where Displayname -eq 'Server - ICMP Ping' | Get-AXNMRulekMaintenanceSchedule

    .Notes
        Author: Jeff Buenting
        Date: 2016 FEB 9
#>

    [CmdletBinding()]
    Param (
        [Parameter(ValueFromPipeline=$True)]
        [PSObject[]]$Rule
    )

    Begin {
        # ----- Get the NM Config object and open the database
        Write-Verbose "Open the Network MOnitoring Database"
        Try {
                $NMConfig = New-Object -ComObject ActiveXperts.NMConfig -ErrorAction Stop 
            }
            catch {
                $EXceptionMessage = $_.Exception.Message
                $ExceptionType = $_.exception.GetType().fullname
                Throw "Get-AXNMRule : Check if ActiveXperts Network Monitoring is installed on $env:ComputerName`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType" 
        }
        $NMConfig.Open()
    }

    Process {
        if ( $Rule ) {
                Write-Verbose "Getting Maintenance Schedule for Rules as well as Global"
                Foreach ( $NMC in $Rule ) {
                    Write-Verbose "Getting Maintenance Schedule for $($NMC.DisplayName)"

                    # ----- Only return the schedules used.  Global (0) or local (255)
                    if ( $NMC.MaintenanceServer -eq 0 ) {
                            Write-Verbose "Getting Global maintenance Schedule"

                            # -----  Get the Maintenance for the Overall tool
                            ($NMConfig.LoadMaintenanceSettings()) -split '\|' | foreach {
                                # ----- Check if Current mainlist is null.  Ignore if it is.  For some reason an empty line is returned if no maintenance schedule is defined.
                                if ( $_ ) {
                                    $MaintSched = $_ | Convert-AXNMMaintScheduletoDate
                                    $MaintSched | Add-Member -MemberType NoteProperty -Name Scope -Value Global
                                    $MaintSched | Add-Member -MemberType NoteProperty -Name RuleName -Value $NMC.DisplayName
                                    Write-Output $MaintSched
                                }
                            }
                        }
                        else {
                            Write-verbose "Local Rule Maintenance Schedule"

                            # ----- Get the Schedule for the NM Check

                            ($NMC.MaintenanceList) -Split '\|' | foreach {
                                # ----- Check if Current mainlist is null.  Ignore if it is.  For some reason an empty line is returned if no maintenance schedule is defined.
                                if ( $_ ) {
                                    $MaintSched = $_ | Convert-AXNMMaintScheduletoDate
                                    $MaintSched | Add-Member -MemberType NoteProperty -Name Scope -Value $NMC.DisplayName
                                    $MaintSched | Add-Member -MemberType NoteProperty -Name RuleName -Value $NMC.DisplayName
                                    Write-Output $MaintSched
                                }
                            }
                    }
                }
            }
            else {
                Write-Verbose "Getting only the Global Maintenance Schedule"
                 # -----  Get the Maintenance for the Overall tool
                ($NMConfig.LoadMaintenanceSettings()) -split '\|' | foreach {
                    # ----- Check if Current mainlist is null.  Ignore if it is.  For some reason an empty line is returned if no maintenance schedule is defined.
                    if ( $_ ) {
                        $MaintSched = $_ | Convert-AXNMMaintScheduletoDate
                        $MaintSched | Add-Member -MemberType NoteProperty -Name Scope -Value Global
                        $MaintSched | Add-Member -MemberType NoteProperty -Name RuleName -Value $NMC.DisplayName
                        Write-Output $MaintSched
                    }
                }
        }
    }

    End {
        Write-Verbose "Closing the Network Monitoring Database"
        $NMConfig.Close()
    }
        
}

#----------------------------------------------------------------------------------

Function Convert-AXNMDatetoMaintenanceSchedule {

<#
    .Synopsis
        Convert real date to an ActiveXperts Network Monitoring Maintenance Schedule,

    .Description
        Convert real date to an ActiveXperts Network Monitoring Maintenance Schedule,  Untility function  Not really needed to be exposed.

    .Parameter MaintDate
        Maintenance date in ActiveXperts NM format

    .Note
        Author: Jeff Buenting
        Date: 2016 FEB 10
#>

    [CmdletBinding()]
    param (
        [String]$MaintenanceSched,

        [Int]$Duration
    )
    
    Try {
            # ----- Convert Date and Time to values ActiveXpert Understands (unix time)
            $DT = $MaintenanceSched -split '@'
        
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
                    $Day = New-TimeSpan -Start "01/01/1970 00:00" -End (Get-Date $MaintenanceSched -ErrorAction Stop).Date -ErrorAction Stop | Select-Object -ExpandProperty TotalSeconds
                    $Day ="d$Day"
                    $T = ( " {0:T}" -f (Get-Date $MaintenanceSched -ErrorAction Stop) )
                }
            }

            # ----- ActiveXpert Net work Monitoring time is measured in Epoch time.  However for some reason, midnight is Jan 3, 1970 midnight.  UTC.  So convert and mesure from those dates to get the time.
            $Time = New-TimeSpan -Start "01/01/1970 00:00" -End ((Get-Date "Saturday, January 3, 1970$T" -ErrorAction Stop).ToUniversalTime()) -ErrorAction Stop | Select-Object -ExpandProperty TotalSeconds
        }
        Catch {
            Throw "Convert-AXNMDatetoMaintenanceSchedule : $($_.Exception.Message)"
    }

    Write-Output "$Day;$Time;$Duration"

}

#----------------------------------------------------------------------------------

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
        [ValidateScript ( {
            # ----- Must be duration larger than one hour
            $_ -ge 1
        } ) ]
        [String]$Duration
    )

    Begin {
        # ----- Don't know why but this date evaluates to back 24 hours.  in ActiveXpert.  To prevent this, adding 24 hours to date.
        $NewMaintSched = Convert-AXNMDatetoMaintenanceSchedule -MaintenanceSched ((Get-Date $MaintenanceSched).AddDays(1)).toString() -Duration $Duration
   
        # ----- Get the NM Config object and open the database
        Write-Verbose "Open the Network MOnitoring Database"
        Try {
                $NMConfig = New-Object -ComObject ActiveXperts.NMConfig -ErrorAction Stop 
            }
            catch {
                $EXceptionMessage = $_.Exception.Message
                $ExceptionType = $_.exception.GetType().fullname
                Throw "Get-AXNMRule : Check if ActiveXperts Network Monitoring is installed on $env:ComputerName`n`n     $ExceptionMessage`n`n     Exception : $ExceptionType"
        }
        $NMConfig.Open()
    }

    Process {
        if ( $Rule ) {
                Write-Verbose "Adding Maintenace schedule for rule: $($Rule.Displayname)"
                Write-Verbose "New Maintenance Schedule: $NewMaintSched"
                # ----- Load Rule from the Database modify and save
                $Node = $NMConfig.LoadNode( $Rule.ID )
                #$Node | FL *
                $Node.Maintenancelist += "|$NewMaintSched"

                # ----- Maintenance Server set to 255 tells the rule to override the Global Rules
                $Node.MaintenanceServer = 255
                $NMConfig.SaveNode( $Node )
                
            }
            else {
                Write-Verbose "Adding Global Maintenance Schedule"
                Write-Verbose "New Maintenance Schedule: $NewMaintSched"
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

#----------------------------------------------------------------------------------

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

                            #----- Set to Global maintenance list if no local ones exist anymore
                            if ( $List -eq $Null ) { $Node.MaintenanceList = 0 }
                                          
                            $NMConfig.SaveNode( $Node )
                        }
                        Else {
                            Throw "Remove-AXNMMaintenanceSchedule : Cannot remove a Global schedule from a Rule. `n$_.Exception.Message"
                    }
                
                }
                elseif ( $M.Scope -ne 'Global' ) {
                    Write-Verbose "No Rule specified.  Taking rule from Schedule object scope"

                    $RuleCheck = Get-AXNMRule | where Displayname -eq $M.Scope

                    Write-Verbose "Rule $($RuleCheck.ID)"


                    # ----- Load Rule from the Database modify and save
                    $Node = $NMConfig.LoadNode( $RuleCheck.ID )
                    
                    # ----- Split the schedule and Loop thru and skip the schedule being deleted
                    $List = $NUll
                    ($Node.Maintenancelist) -split '\|' | Foreach {
                        if ( $_ ) {
                            Write-Verbose "Converting $_"
                            $MS = $_ | Convert-AXNMMaintScheduletoDate
                                 
                            if ( ($MS.Date -ne $M.date) -and ($MS.Duration -ne $M.Duration ) ) {
                                Write-Verbose "     Keeping $MS"
                                # ----- Add divider if list is not null
                                if ( $List -ne $Null ) { $List += '|' }
                                $List += "$_"
                            }
                        }
                    }
                           
                    $Node.MaintenanceList = $List
                    
                    #----- Set to Global maintenance list if no local ones exist anymore
                    if ( $List -eq $Null ) { $Node.MaintenanceList = 0 }
                                          
                    $NMConfig.SaveNode( $Node ) 
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

#----------------------------------------------------------------------------------

#----------------------------------------------------------------------------------