Import-Module activedirectory

$days = "-1"
$date = Get-Date -format D
$servers = "sb-dc01","sb-dc02","s4b-wsus"
$allid = "4","5","6"


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


function EnviaEmailUser ($destino,$mensagem,$assunto) 
{
$email = New-Object System.Net.Mail.MailMessage("infra@s4bdigital.net",$destino)
$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
$email.Subject = $assunto
$email.Body = $mensagem
$email.Priority = 2
$email.IsBodyHtml = $True
$smtp.Send($email)
}



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

$report = $table |Select-Object LogName,EventID,Server,Level,Message,TimeCreated |Format-Table | Out-String

EnviaEmailUser -Destino $emailTo -Mensagem $report -Assunto "Report - Backup Full Windows Servers - SBD.CORP"
