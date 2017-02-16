Import-Module activedirectory

$days = "-1"
$date = Get-Date -format D
$servers = "sb-dc01","sb-dc02","s4b-wsus"
$allid = "4","5","6"
$outfile = "e:\usr\util\scripts\logs\Rel_Bkp_SystemState_Servers.html"
$img = "e:\img\s4bdigital.jpg"
$css = "e:\usr\util\scripts\HtmlReports.css"

$table = New-Object system.Data.DataTable "TableSample"
$col1 = New-Object system.Data.DataColumn LogName ,([string])
$col2 = New-Object system.Data.DataColumn EventID ,([string])
$col3 = New-Object system.Data.DataColumn Server ,([string])
$col4 = New-Object system.Data.DataColumn Message ,([string])
$col5 = New-Object system.Data.DataColumn TimeCreated ,([string])
$col6 = New-Object system.Data.DataColumn Level ,([string])
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)



$SrvPassword = ConvertTo-SecureString "$($ENV:SrvPassword)" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ("$ENV:SrvUser", $SrvPassword)



foreach( $allservers in $servers ) {

	invoke-command -Computername $allservers -Credential $Credential -scriptblock { 
    
        foreach ( $allid in $id ) {

			$EventLog = Get-WinEvent -FilterHashtable @{ logname="Microsoft-Windows-Backup"; id="$allid"; StartTime=(get-date).addDays($days)} -Computername $allservers
					
                $row = $table.NewRow();
                $row.LogName = "Microsft-Windows-Backup"; 
		        $row.Server = $allservers;
                $row.Level = "Success";
		        $row.TimeCreated = [string]$Eventlog.TimeCreated;
		
				switch -regex ($EventLog.id) {
		                                                   
					"[4]" {  
                  
                        $row.EventID = [string]$EventLog.id;
                        $row.Message = [string]$Eventlog.Message;
		          	    
						break;
					} 
		
					"[5]"{  
				
                    	$row.EventID = [string]$EventLog.id;
		       			$row.Message = [string]$Eventlog.Message;
						break;
					}
				
					"[7]"{  
				
                        $row.EventID = [string]$EventLog.id;
		       		    $row.Message = [string]$Eventlog.Message;
						break;
					}

				}
                $table.Rows.Add($row)

		    }

	 	
    	} 
   

 
}  

$log = $table |Select-Object LogName,EventID,Server,Level,Message,TimeCreated |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Report Backup Full - Windows Servers</h4>" | Out-String 

$report = ConvertTo-Html -CSSUri $css -Title "Report Backup Full - Windows Servers" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log" |Out-String
$report | Out-File $outfile | Out-String

