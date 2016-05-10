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

#>

    [CmdletBinding()]
    Param (
        [ValidateSet( 'ALL','CHECKTYPE_ADOSQL','CHECKTYPE_CPU','CHECKTYPE_DIRSIZE','CHECKTYPE_DISKS','CHECKTYPE_DISKSPACE','CHECKTYPE_DNS','CHECKTYPE_DOOR','CHECKTYPE_EVENTLOG','CHECKTYPE_FILE','CHECKTYPE_FLOPPY','CHECKTYPE_FOLDER','CHECKTYPE_FTP','CHECKTYPE_HTTP','CHECKTYPE_HUMIDITY','CHECKTYPE_ICMP','CHECKTYPE_IMAP','CHECKTYPE_LIGHT','CHECKTYPE_MEMORY','CHECKTYPE_MOTION','CHECKTYPE_MSMQ','CHECKTYPE_MSTSE','CHECKTYPE_NNTP','CHECKTYPE_NTP','CHECKTYPE_ODBC','CHECKTYPE_ORACLE','CHECKTYPE_POP3','CHECKTYPE_POWER','CHECKTYPE_PRINTER','CHECKTYPE_PROCESS','CHECKTYPE_REGISTRY','CHECKTYPE_RESISTANCE','CHECKTYPE_RSH','CHECKTYPE_SERVICE','CHECKTYPE_SMOKE','CHECKTYPE_SMTP','CHECKTYPE_SNMPGET','CHECKTYPE_SNMPTRAPRECEIVE','CHECKTYPE_SSH','CHECKTYPE_SWITCHNC','CHECKTYPE_SWITCHNO','CHECKTYPE_TCPIP','CHECKTYPE_TEMPERATURE','CHECKTYPE_UNDEFINED','CHECKTYPE_VBSCRIPT','CHECKTYPE_WETNESS' )]
        [String[]]$CheckType = 'ALL'
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
    
    foreach ( $Check in $CheckType ) {
        Write-Verbose "Getting the following Checks:"
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
   
    Write-Verbose "Closing the Network Monitoring Database"
    $NMConfig.Close()
}

$Cred = Get-Credential

$Check = Get-AXNMRule -ComputerName RWVA-TS1 -Credential $Cred  -Verbose 

#| where Displayname -eq 'Rwva-ts1 - ICMP Ping' #| FT Displayname, Maintenancelist










