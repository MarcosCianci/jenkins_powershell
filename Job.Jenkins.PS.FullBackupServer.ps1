###########################################################
# PowerShell Windows 
# Script: Job.Jenkins.PS.FullBackupServer.ps1
#
# Autor: Marcos Cianci
# Data: 22/01/2013
#
# Atualizações:
#               06/04/2015 - Marcos Cianci
#               16/02/2017 - Marcos Cianci   
#    
#
# Versão 1.0:  Script verifica a execução do bkp 
#				do Windows Server 2012 R2.	
#
# Versão 2.0 : Script verifica nos logs dos servidores:
#              sb-dc01, sb-dc02,s4b-wsus,s4b-bi o status      
#              dos backups realizados nos servidores,
#              envia por email relatorio gerado em HTML.
#
# Versão 3.0: Script integrado ao Jenkins
#
###########################################################

### MODULOS ###
Import-Module activedirectory

### VARIÁVEIS ###
$days = "-1"
$date = Get-Date -format D
$servers = "sb-dc01","sb-dc02","s4b-wsus","s4b-acesso","s4b-bi"
$id = "4","5","7"
$outfile = "e:\usr\util\scripts\logs\Rel_Bkp_SystemState_Servers.html"
$img = "\\sb-dc01\img\s4bdigital.jpg"
$css = "e:\usr\util\scripts\HtmlReports.css"

### TABELA ###
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

### CREDENCIAIS WINDOWS ###
$SrvPassword = ConvertTo-SecureString "$($ENV:SrvPassword)" -AsPlainText -Force
$Credential = New-Object System.Management.Automation.PSCredential ("$ENV:SrvUser", $SrvPassword)

### STARTING ###

foreach( $allservers in $servers ) {

	invoke-command -Computername $allservers -Credential $Credential -scriptblock { 
    
        foreach ( $allid in $id ) {

			$EventLog = Get-WinEvent -FilterHashtable @{ logname="Microsoft-Windows-Backup"; id="$allid"; StartTime=(get-date).addDays($days)} -Computername $allservers
					
                $row = $table.NewRow();
                $row.LogName = "Microsft-Windows-Backup"; 
		        $row.Server = $allservers;
                $row.TimeCreated = [string]$Eventlog.TimeCreated;
		
				switch -regex ($EventLog.id) {
		                                                   
					"[4]" {  
                  
                        $row.EventID = [string]$EventLog.id;
                        $row.Message = [string]$Eventlog.Message;
                        $row.Level = "Success";
		          	    
						break;
					} 
		
					"[5]"{  
				
                    	$row.EventID = [string]$EventLog.id;
		       			$row.Message = [string]$Eventlog.Message;
                        $row.Level = "Error"; 
						break;
					}
				
					"[7]"{  
				
                        $row.EventID = [string]$EventLog.id;
		       		    $row.Message = [string]$Eventlog.Message;
                        $row.Level = "Warning";
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

### FINISH ###