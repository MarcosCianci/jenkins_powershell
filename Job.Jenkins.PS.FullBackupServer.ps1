Import-Module activedirectory

$days = "-1"
$date = Get-Date -format D
$servers = "sb-dc01","sb-dc02","s4b-wsus"
$id = "4","5","7"
$ErrorActionPreference = 'Stop'


$SrvPassword = ConvertTo-SecureString "$($ENV:SrvPassword)" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ("$ENV:SrvUser", $SrvPassword)

foreach( $allservers in $servers ) {

	invoke-command -Computername $allservers -Credential $Credential -scriptblock { 

		foreach ( $allid in $id ) {
			$EventLog = Get-WinEvent -FilterHashtable @{ logname="Microsoft-Windows-Backup"; id="$allid"; StartTime=(get-date).addDays($days)} -computerName $allservers
					
				switch -regex ($EventLog.id) {
		
					"[4]" {  
								$log = $EventLog 
								break;
					} 
		
					"[5]"{  
					
						$log = $EventLog
						break;
					}
				
					"[7]"{  
					
						$log = $EventLog
						break;
					}
				}
		}
	 	
    	} 
   

write-output $EventLog   
}  


