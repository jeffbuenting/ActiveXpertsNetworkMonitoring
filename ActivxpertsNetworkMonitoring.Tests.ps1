$ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path

$ModuleName = $ModulePath | Split-Path -Leaf

# ----- Remove and then import the module.  This is so any new changes are imported.
Get-Module -Name $ModuleName -All | Remove-Module -Force -Verbose

Import-Module "$ModulePath\$ModuleName.PSD1" -Force -ErrorAction Stop -Scope Global -Verbose

#-------------------------------------------------------------------------------------

Write-Output "`n`n"

Describe "AxtiveXperts : Get-AXNMRule" {
    Context "Help" {
        $H = Help Get-AXNMRule -Full
        
        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 
    
    Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
        $OBJ = New-Object -TypeName PSObject -Property (@{
                ConfigDatabase = "Connection String"
                LastError = 0
            })
        
        $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name FindFirstNode -Value {
            

            $N = New-Object -TypeName PSObject -Property (@{
                DisplayName = 'Test Rule'
                ID = 46789
                Type = 33
                MaintenanceServer = 0
                MaintenanceList = "e0010000;270000"
            })

            Write-Output $N

        }
        $OBJ | Add-Member -MemberType ScriptMethod -Name FindNextNode -Value {

            $Nodes = $N

            $N2 = New-Object -TypeName PSObject -Property (@{
                DisplayName = 'Test Rule 2'
                ID = 46783
                Type = 25
                MaintenanceServer = 0
                MaintenanceList = "e0010000;270000"
            })

            $Nodes += $N2

            Write-Output $N

        }
        $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
            # ----- Should return a string of Maintenance windows separated by |
            Write-Output "e0100000;270000;1|e0010000;27000;1|e0000100;270000;1"
        }
        $OBJ | Add-Member -memberType ScriptMethod -Name LoadNode -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name SaveMaintenanceSettings -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name SaveNode -Value { }

        Write-Output $OBJ
    } 

    Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConstants' } -MockWith {
        $C = New-Object -TypeName PSObject -Property (@{
            CHECKTYPE_UNDEFINED       = 0
            CHECKTYPE_FOLDER          = 1
            CHECKTYPE_ADOSQL          = 111
            CHECKTYPE_CPU             = 130
            CHECKTYPE_DIRSIZE         = 140
            CHECKTYPE_DISKS           = 150
            CHECKTYPE_DISKSPACE_WMI   = 20
            CHECKTYPE_DISKSPACE_NAS   = 21
            CHECKTYPE_DNS             = 36
            CHECKTYPE_EVENTLOG        = 90
            CHECKTYPE_FILE            = 60
            CHECKTYPE_FLOPPY          = 160
            CHECKTYPE_FTP             = 34
            CHECKTYPE_HTTP            = 33
            CHECKTYPE_HUMIDITY        = 200
            CHECKTYPE_ICMP            = 10
            CHECKTYPE_IMAP            = 39
            CHECKTYPE_MEMORY          = 170
            CHECKTYPE_MSTSE           = 120
            CHECKTYPE_NNTP            = 38
            CHECKTYPE_NTP             = 40
            CHECKTYPE_ODBC            = 110
            CHECKTYPE_ORACLE          = 112
            CHECKTYPE_POP3            = 31
            CHECKTYPE_PRINTER         = 180
            CHECKTYPE_PROCESS         = 190
            CHECKTYPE_REGISTRY        = 61
            CHECKTYPE_RSH             = 35
            CHECKTYPE_SSH             = 43
            CHECKTYPE_SERVICE         = 50
            CHECKTYPE_SMTP            = 32
            CHECKTYPE_SNMPGET         = 37
            CHECKTYPE_SNMPTRAPRECEIVE = 42
            CHECKTYPE_TEMPERATURE     = 41
            CHECKTYPE_TCPIP           = 30
            CHECKTYPE_VBSCRIPT        = 100
            CHECKTYPE_WETNESS         = 210
            CHECKTYPE_POWER           = 201
            CHECKTYPE_LIGHT           = 202
            CHECKTYPE_MOTION          = 203
            CHECKTYPE_SMOKE           = 204
            CHECKTYPE_DOOR            = 205
            CHECKTYPE_RESISTANCE      = 206
            CHECKTYPE_SWITCHNC        = 207
            CHECKTYPE_SWITCHNO        = 208
            CHECKTYPE_MSMQ            = 250
            NODEID_UNDEFINED          = 0
            NODEID_ROOT               = 1
            NODEID_DEFAULTSETTINGS    = 2
            NODEID_USERBASE           = 10000
            RESULT_UNCERTAIN          = 8
            RESULT_SUCCESS            = 1
            RESULT_ERROR              = 2
            RESULT_FAILURE            = 3
            RESULT_MAINTENANCE        = 4
            RESULT_ONHOLD             = 5
            RESULT_DEPENDEE_ERROR     = 6
            RESULT_DEPENDEE_FAILURE   = 7
            RESULT_NOTPROCESSED       = 0
            DBTYPE_MDB                = 1
            DBTYPE_SDF                = 3
            DBTYPE_MSSQL              = 4
            DBTYPE_MYSQL              = 5
            DBTYPE_ORACLE             = 6
        })

        Write-Output $C
    }

    Context Execution {

        It "Should throw an error if there is a problem creating an ActiveXpert COM Object" {
             Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -MockWith { Throw "Test Error" }
            { Get-AXNMRule } | Should Throw
        } 

        It "CheckType : Returns all rules of a specific type" {

            Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
                $OBJ = New-Object -TypeName PSObject -Property (@{
                    ConfigDatabase = "Connection String"
                    LastError = 0
                })
        
                $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
                $OBJ | Add-Member -MemberType ScriptMethod -Name FindFirstNode -Value {
            

                    $N = New-Object -TypeName PSObject -Property (@{
                        DisplayName = 'Test Rule'
                        ID = 46789
                        Type = 33
                        MaintenanceServer = 0
                        MaintenanceList = "e0010000;270000"
                    })

                    Write-Output $N

                }
                $OBJ | Add-Member -MemberType ScriptMethod -Name FindNextNode -Value {

                    $Nodes = $N

                    $N2 = New-Object -TypeName PSObject -Property (@{
                        DisplayName = 'Test Rule 2'
                        ID = 46783
                        Type = 25
                        MaintenanceServer = 0
                        MaintenanceList = "e0010000;270000"
                    })

                    $Nodes += $N2

                    Write-Output $N

                }
                $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
                    # ----- Should return a string of Maintenance windows separated by |
                    Write-Output "e0100000;270000;1|e0010000;27000;1|e0000100;270000;1"
                }
                $OBJ | Add-Member -memberType ScriptMethod -Name LoadNode -Value { }
                $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
                $OBJ | Add-Member -MemberType ScriptMethod -Name SaveMaintenanceSettings -Value { }
                $OBJ | Add-Member -MemberType ScriptMethod -Name SaveNode -Value { }

                Write-Output $OBJ
            }     

            Get-AXNMRule #| Should Be 2

            # | Measure-Object | Select-Object -ExpandProperty Count
        } 

        It "ID : Returns the specific Rule ID" {

        } -Pending
    } 

    Context Output {

        It "Should retrieve a custom Rule PSObject" {
            

            Get-AXNMRule  | Should BeOfType PSObject
        } -Pending

    }
}

#-------------------------------------------------------------------------------------

Write-Output "`n`n"

InModuleScope ActiveXpertsNetworkMonitoring {

    Describe "AxtiveXperts : Convert-AXNMMaintScheduletoDate" {
       Context "Help" {
            $H = Help Convert-AXNMMaintScheduletoDate -Full
        
            # ----- Help Tests
            It "has Synopsis Help Section" {
                $H.Synopsis | Should Not BeNullorEmpty
            }

            It "has Description Help Section" {
                $H.Description | Should Not BeNullorEmpty
            }

            It "has Parameters Help Section" {
                $H.Parameters | Should Not BeNullorEmpty
            }

            # Examples - Remarks (small description that comes with the example)
            foreach ($Example in $H.examples.example)
            {
                it "Example - Remarks on $($Example.Title)"{
                    $Example.remarks | Should not BeNullOrEmpty
                }
            }

            It "has Notes Help Section" {
                $H.alertSet | Should Not BeNullorEmpty
            }
        }  

        Context Output {
            It "returns custom maintenance schedule object" {
                Convert-AXNMMaintScheduletoDate -MaintDate "e0100000;270000;1" | Should beoftype PSObject
            }

            It "should have a duration of 1"{
                (Convert-AXNMMaintScheduletoDate -MaintDate "e0100000;270000;1").Duration | Should be 1
            }

            It "should have a Date of Every Monday @ 10:00:00 PM"{
                (Convert-AXNMMaintScheduletoDate -MaintDate "e0100000;270000;1").Date | Should be "Every Monday @ 10:00:00 PM"
            }
        }
    }

    #-------------------------------------------------------------------------------------

    Write-Output "`n`n"

    Describe "AxtiveXperts : Convert-AXNMDatetoMaintenanceSchedule" {
       Context "Help" {
            $H = Help Convert-AXNMDatetoMaintenanceSchedule -Full
        
            # ----- Help Tests
            It "has Synopsis Help Section" {
                $H.Synopsis | Should Not BeNullorEmpty
            }

            It "has Description Help Section" {
                $H.Description | Should Not BeNullorEmpty
            }

            It "has Parameters Help Section" {
                $H.Parameters | Should Not BeNullorEmpty
            }

            # Examples - Remarks (small description that comes with the example)
            foreach ($Example in $H.examples.example)
            {
                it "Example - Remarks on $($Example.Title)"{
                    $Example.remarks | Should not BeNullOrEmpty
                }
            }

            It "has Notes Help Section" {
                $H.alertSet | Should Not BeNullorEmpty
            }
        } 

        Context Execution {
            It "Throws an error if invalid input is given" {
                { Convert-AXNMDatetoMaintenanceSchedule -MaintenanceSched "funday @ 12:00 pm" -Duration 2 } | Should Throw
            }
        }
        
        Context Output {
            It "returns a string in the format Day;Time;Duration" {
                Convert-AXNMDatetoMaintenanceSchedule -MaintenanceSched "every tuesday @ 12:00 pm" -Duration 2 | Should be 'e0010000;234000;2'
            }
        } 
    }
}

#-------------------------------------------------------------------------------------

Write-Output "`n`n"


Describe "AxtiveXperts : New-AXNMMaintenanceSchedule" {
   
    # ----- Get Function Help
    # ----- Pester to test Comment based help
    # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html

    Context "Help" {
        $H = Help New-AXNMMaintenanceSchedule -Full
        
        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples
        it "Example - Count should be greater than 0"{
            $H.examples.example.code.count | Should BeGreaterthan 0
        }

        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 

    $Rule = New-Object -TypeName PSObject -Property (@{
        DisplayName = 'Test Rule'
        MaintenanceServer = 0
        MaintenanceList = "e0010000;270000"
    })    

    Context Execution {

        Mock -CommandName New-Object -MockWith { Throw "Test Error" }

        It "Should throw an error if there is a problem creating an ActiveXpert COM Object" {
            { New-AXNMMaintenanceSchedule -Rule $Rule -MaintenanceSched 'every tuesday @ 1:00 am'-Duration 1 } | Should Throw
        } 


    }

    Context Output {

        Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
            $OBJ = New-Object -TypeName PSObject -Property (@{
                    ConfigDatabase = "Connection String"
                    LastError = 0
                })
        
            $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
            $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
                # ----- Should return a string of Maintenance windows separated by |
                Write-Output "e0100000;270000;1|e0010000;27000;1|e0000100;270000;1"
            }
            $OBJ | Add-Member -memberType ScriptMethod -Name LoadNode -Value {
                $Node = New-Object -TypeName PSObject -Property (@{
                    DisplayName = 'Test Rule'
                    MaintenanceServer = 0
                    MaintenanceList = "e0010000;270000"
                })

                Write-Output $Node
            }
            $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
            $OBJ | Add-Member -MemberType ScriptMethod -Name SaveMaintenanceSettings -Value { }
            $OBJ | Add-Member -MemberType ScriptMethod -Name SaveNode -Value { }

            Write-Output $OBJ
        }



        # ----- Maintenance Windows for Rule

        It "Rule : Does not return anything by default" {
            New-AXNMMaintenanceSchedule -Rule $Rule -MaintenanceSched "every tuesday @ 10:00 am" -Duration 4 | Should BeNullOrEmpty
        }

        It "Rule : Returns New maintenance object if Passthru specified" {
            New-AXNMMaintenanceSchedule -Rule $Rule -MaintenanceSched "every tuesday @ 10:00 am" -Duration 4 -PassThru | Should beoftype PSObject
        } 

        # ----- Maintenance Window for Global

        It "Global Maint Sched : Does not return anything by default" {
            New-AXNMMaintenanceSchedule -MaintenanceSched "every tuesday @ 10:00 am" -Duration 4 | Should BeNullOrEmpty
        } 

        It "Global Maint Sched : returns New maintenance object if Passthru specified" {
            New-AXNMMaintenanceSchedule -MaintenanceSched "every tuesday @ 10:00 am" -Duration 4 -PassThru | Should beoftype PSObject
        }
    }
}

#-------------------------------------------------------------------------------------

Write-Output "`n`n"


Describe "AxtiveXperts : Get-AXNMMaintenanceSchedule" {
   
    # ----- Get Function Help
    # ----- Pester to test Comment based help
    # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html

    Context "Help" {
        $H = Help Get-AXNMMaintenanceSchedule -Full
        
        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples
        it "Example - Count should be greater than 0" {
            $H.examples.example.code.count | Should BeGreaterthan 0
        }

        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 


    Context Output {

     #   Mock -CommandName Convert-AXNMMaintScheduletoDate -ModuleName ActiveXpertsNetworkMonitoring -MockWith {
     #       $Obj = New-Object -TypeName PSObject -Property (@{  
     #           Date = "Every Monday @ 10:00:00 PM"
     #           Duration = 3
     #       })
     #   
     #       Return $Obj
     #   }

        # ----- Test Retrieving Global Maintenance Schedules
        Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
            $OBJ = New-Object -TypeName PSObject -Property (@{
                    ConfigDatabase = "Connection String"
                    LastError = 0
                })
        
            $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
            $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
                # ----- Should return a string of Maintenance windows separated by |
                Write-Output "e0100000;270000;1|e0010000;27000;1|e0000100;270000;1"
            }
            $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
            

            Write-Output $OBJ
        }

        It "Global Schedules : Should return 3 global windows" {
           Get-AXNMMaintenanceSchedule | Measure-Object | Select-Object -ExpandProperty Count | Should beexactly 3
        }

        It "Global Schedules : Should Return a custom Maintenance schedule Object" {
            Get-AXNMMaintenanceSchedule  | Should beoftype PSObject
        }

        # ----- Test if Rule has nothing scheduled

        It "Rule empty schdule : Should return 3 global windows" {
           Get-AXNMMaintenanceSchedule | Measure-Object | Select-Object -ExpandProperty Count | Should beexactly 3
        }

        It "Rule empty schdule : Should Return a custom Maintenance schedule Object" {
            Get-AXNMMaintenanceSchedule  | Should beoftype PSObject
        }

        # ----- Test REtrieving Maintenance Schedule for a Rule
        

        It "Monitor Maintenance Schedules : Should return 1 global windows" {
            Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
                $OBJ = New-Object -TypeName PSObject -Property (@{
                        ConfigDatabase = "Connection String"
                        LastError = 0
                    })
        
                $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
                $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
                    # ----- Should return a string of Maintenance windows separated by |
                    Write-Output "e0100000;270000"
                }
                $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
            

                Write-Output $OBJ
            }

            $Rule = New-Object -TypeName PSObject -Property (@{
                DisplayName = 'Test Rule'
                MaintenanceServer = 0
                MaintenanceList = "e0010000;270000"
            })

           Get-AXNMMaintenanceSchedule -Rule $Rule | Measure-Object | Select-Object -ExpandProperty Count | Should beexactly 1
        }

        It "Monitor Maintenance Schedules  : Should Return a custom Maintenance schedule Object" {
            Get-AXNMMaintenanceSchedule -Rule $Rule  | Should beoftype PSObject
        } 
        
        # ----- Test if Rule has nothing scheduled

        
    }
}

#-------------------------------------------------------------------------------------

Write-Output "`n`n"


Describe "AxtiveXperts : Remove-AXNMMaintenanceSchedule" {
   
    # ----- Get Function Help
    # ----- Pester to test Comment based help
    # ----- http://www.lazywinadmin.com/2016/05/using-pester-to-test-your-comment-based.html

    Context "Help" {
        $H = Help Remove-AXNMMaintenanceSchedule -Full
        
        # ----- Help Tests
        It "has Synopsis Help Section" {
            $H.Synopsis | Should Not BeNullorEmpty
        }

        It "has Description Help Section" {
            $H.Description | Should Not BeNullorEmpty
        }

        It "has Parameters Help Section" {
            $H.Parameters | Should Not BeNullorEmpty
        }

        # Examples
        it "Example - Count should be greater than 0" {
            $H.examples.example.code.count | Should BeGreaterthan 0
        }

        # Examples - Remarks (small description that comes with the example)
        foreach ($Example in $H.examples.example)
        {
            it "Example - Remarks on $($Example.Title)"{
                $Example.remarks | Should not BeNullOrEmpty
            }
        }

        It "has Notes Help Section" {
            $H.alertSet | Should Not BeNullorEmpty
        }
    } 

    $Rule = New-Object -TypeName PSObject -Property (@{
        DisplayName = 'Test Rule'
        MaintenanceServer = 0
        MaintenanceList = "e0010000;270000"
    }) 
    
     Mock -CommandName New-Object -ModuleName ActiveXpertsNetworkMonitoring -ParameterFilter { $ComObject -eq 'ActiveXperts.NMConfig' } -MockWith {
        $OBJ = New-Object -TypeName PSObject -Property (@{
                ConfigDatabase = "Connection String"
                LastError = 0
            })
        
        $OBJ | Add-Member -MemberType ScriptMethod -Name Close -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name LoadMaintenanceSettings -Value {
            # ----- Should return a string of Maintenance windows separated by |
            Write-Output "e0100000;270000;1|e0010000;27000;1|e0000100;270000;1"
        }
        $OBJ | Add-Member -memberType ScriptMethod -Name LoadNode -Value {
            $Node = New-Object -TypeName PSObject -Property (@{
                DisplayName = 'Test Rule'
                MaintenanceServer = 0
                MaintenanceList = "e0010000;270000"
            })

            Write-Output $Node
        }
        $OBJ | Add-Member -MemberType ScriptMethod -Name Open -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name SaveMaintenanceSettings -Value { }
        $OBJ | Add-Member -MemberType ScriptMethod -Name SaveNode -Value { }

        Write-Output $OBJ
    }   

    Context Execution {
       
        It "Rule : Throws an error when the Maintanance Schedule to be removed is global and a Rule is included" {
            
            $MaintSched = New-Object -TypeName PSObject -Property (@{
                Duration = 12
                Date = 'Every Monday @ 10:00 AM'
                Scope = 'Global'
                RuleName = ''
            })

            {Remove-AXNMMaintenanceSchedule -MaintenanceSchedule $MaintSched -Rule $Rule} | Should Throw
        } 

        it "Rule : Does not throw and error " {
            $MaintSched = New-Object -TypeName PSObject -Property (@{
                Duration = 12
                Date = 'Every Monday @ 10:00 AM'
                Scope = 'Test Rule'
                RuleName = 'Test Rule'
            })

            {Remove-AXNMMaintenanceSchedule -MaintenanceSchedule $MaintSched -Rule $Rule} | Should Not Throw
        } 

        It "No Rule Specified : Does not throw an error" {
            Mock -CommandName Get-AXNMRule -MockWith {
                $Rule = New-Object -TypeName PSObject -Property (@{
                    DisplayName = 'Test Rule'
                    ID = '17435'
                    MaintenanceServer = 0
                    MaintenanceList = "e0010000;270000"
                })
                
                Write-Object $Rule 
            }

            $MaintSched = New-Object -TypeName PSObject -Property (@{
                Duration = 12
                Date = 'Every Monday @ 10:00 AM'
                Scope = 'Test Rule'
                RuleName = 'Test Rule'
            })

            {Remove-AXNMMaintenanceSchedule -MaintenanceSchedule $MaintSched -verbose } | Should Not Throw
        } 

        It "Global : Does not throw an error" {

        } -Pending

    }

       
}