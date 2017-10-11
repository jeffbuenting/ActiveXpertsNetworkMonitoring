$ModulePath = Split-Path -Parent $MyInvocation.MyCommand.Path

$ModuleName = $ModulePath | Split-Path -Leaf

# ----- Remove and then import the module.  This is so any new changes are imported.
Get-Module -Name $ModuleName -All | Remove-Module -Force -Verbose

Import-Module "$ModulePath\$ModuleName.PSD1" -Force -ErrorAction Stop -Scope Global -Verbose

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

    Context Execution {
        It "Should throw an error if there is a problem creating an ActiveXpert COM Object" {
        } -Pending


    }

    Context Output {
        It "Does not return anything by default" {

        } -Pending

        It "returns New maintenance object if Passthru specified" {

        } -Pending
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
