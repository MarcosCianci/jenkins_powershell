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

$outfile = "e:\usr\util\scripts\logs\Rel_Bkp_SystemState_Servers.html"
$img = "e:\img\s4bdigital.jpg"
$css = "e:\usr\util\scripts\HtmlReports.css"

### CREDENCIAIS WINDOWS ###


$securepassword = ConvertTo-SecureString -String $env:SrvPassword -AsPlainText -Force 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $env:SrvUser, $securepassword 


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


### STARTING ###

foreach( $allservers in $servers ) {

	    
     $EventLog_success = Invoke-Command -Computername $allservers -ScriptBlock { Get-WinEvent -FilterHashtable @{logname="Microsoft-Windows-Backup"; id="4"; StartTime=(get-date).addDays(-1)}} -Credential $cred
	 
     $EventLog_success | foreach {   
                $row = $table.NewRow();
                $row.LogName = "Microsof-Windows-Backup"
                $row.EventID = [string] $EventLog_success.id;
                $row.Server = [string] $EventLog_success.PSComputerName;
                $row.Message = [string]$EventLog_success.Message;
                $row.Level = "Success";
                $row.TimeCreated = [string]$EventLog_success.TimeCreated;
                $table.Rows.Add($row) 
    }

    $EventLog_error = Invoke-Command -Computername $allservers -ScriptBlock { Get-WinEvent -FilterHashtable @{ logname="Microsoft-Windows-Backup"; id="5"; StartTime=(get-date).addDays(-1)}} -Credential $cred

	$EventLog_error | foreach {
                $row = $table.NewRow();
                $row.LogName = "Microsof-Windows-Backup"
                $row.EventID = [string] $EventLog_error.id;
                $row.Server = [string] $EventLog_error.PSComputerName; 
                $row.Message = [string] $EventLog_error.Message;
                $row.Level = "Error";
                $row.TimeCreated = $EventLog_error.TimeCreated;
                $table.Rows.Add($row) 
   }
           

    $EventLog_war = Invoke-Command -Computername $allservers -ScriptBlock { Get-WinEvent -FilterHashtable @{ logname="Microsoft-Windows-Backup"; id="7"; StartTime=(get-date).addDays(-1)}} -Credential $cred

	$EventLog_war | foreach { 
                $row = $table.NewRow();
                $row.LogName = "Microsof-Windows-Backup"
                $row.EventID = [string] $EventLog_war.id;
                $row.Server = [string] $EventLog_error.PSComputerName;
                $row.Message = [string]$EventLog_war.Message;
                $row.Level = "Warning";
                $row.TimeCreated = [string]$EventLog_war.TimeCreated;
                $table.Rows.Add($row) 
        
   }
		    
}  

$log = $table |Select-Object LogName,EventID,Server,Level,Message,TimeCreated |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Report Backup Full - Windows Servers</h4>" | Out-String 

$report = ConvertTo-Html -CSSUri $css -Title "Report Backup Full - Windows Servers" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log" |Out-String
$report | Out-File $outfile | Out-String

### FINISH ###