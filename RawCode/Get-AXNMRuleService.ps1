import-module C:\Scripts\ActiveXpertsNetworkMonitoring\ActiveXpertsNetworkMonitoring.psd1 -force

$Rule = Get-AXNMRule -id 10155

Get-Service -ComputerName $Rule.CheckServer -Name $Rule.Checkparam1