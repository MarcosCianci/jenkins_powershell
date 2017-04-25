###############################################################
# Windows PowerShell 
# 
# Script: Job.Jenkins.PS.EventLog_Error_Servers.ps1
# Autor: Marcos Cianci
# Data: 03/03/2017
# 
#
# Atualização: 
#  
#
# Versão 1.0: Script tem a funcionalidade de verificar os logs de 
#              erros nos servidores windows servers.
#
#
##############################################################

### MODULE ###
Import-Module activedirectory

### VARIABLE GLOBAL ###

$hours = "6"
$date = (Get-Date) - (New-TimeSpan -Hours $hours)
$emailTo = "suporte@s4bdigital.net"
$servers = "SB-DC01","SB-DC02","S4B-WSUS","S4B-ACESSO","S4B-BI","RMNEW"
$Subject="Alerta - Relatório Eventlog Windows Servers - SBD.CORP"
$img = "E:\img\s4bdigital.jpg"
$css = "e:\usr\util\scripts\HtmlReports.css"
$outfile = "e:\usr\util\scripts\logs\Rel_EventLog_Windows_Servers.html"
$logname = "system","application"

### TABLES REPORT ###

$table = New-Object system.Data.DataTable "TableDeleted"
$col1 = New-Object system.Data.DataColumn Level ,([string])
$col2 = New-Object system.Data.DataColumn Message ,([string])
$col3 = New-Object system.Data.DataColumn Id ,([string])
$col4 = New-Object system.Data.DataColumn TimeCreated ,([string])
$col5 = New-Object system.Data.DataColumn Server ,([string])
$col6 = New-Object system.Data.DataColumn Alerta ,([string])
$col7 = New-Object system.Data.DataColumn Logname ,([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)


### CREDENCIAIS WINDOWS ###

$securepassword = ConvertTo-SecureString -String $env:SrvPassword -AsPlainText -Force 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $env:SrvUser, $securepassword 


### FUNCTION ###
## Envio E-mail Suporte ## 

function EnviaEmailUser ($destino,$mensagem,$assunto){
		$email = New-Object System.Net.Mail.MailMessage("infra@s4bdigital.net",$destino)
		$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
        $email.Subject = $assunto
		$email.Body = $mensagem
		$email.Priority = 2
        $email.IsBodyHtml = $True
		$smtp.Send($email)
}


### START ###


################################################
## EVENTO LOG ERROR  - SYSTEM AND APPLICATION ##
################################################

foreach ( $allserver in $servers ){
 

 foreach ( $logname in $logname ){

    $EventError = Invoke-Command -ComputerName $allserver -ScriptBlock { Get-WinEvent -FilterHashtable @{logname='application','system';starttime=[datetime]::today;level=2 }} -Credential $cred
    $EventError | foreach {

    $row=$table.NewRow()
    $row.Alerta = "Alert - EventLog Error"
    $row.Logname = $logname 
    $row.Level = $EventLog.Level
    $row.Message = $EventLog.Message 
    $row.Id = $Eventlog.Id
    $row.Server = $allserver
    $row.TimeCreated = $EventLog.TimeCreated
    $table.Rows.Add($row) 
    
    }
  }
}


$log = $table |Select-Object Alerta,Level,Id,Message,Server,TimeCreated  | Sort-Object Server |ConvertTo-Html -Fragment -As Table -PreContent "<h4> Report Alert EventLog Windows Servers - SBD.CORP </h4>" | Out-String 

$report = ConvertTo-Html -CSSUri $css -Title "Report Alert EventLog Windows Servers - SBD.CORP " -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log" |Out-String
$report | Out-File $outfile | Out-String

$emailTo = "marcos.cianci@s4bdigital.net"
$subject = "Report Alert EventLog Windows Servers - SBD.CORP"
$body_rel = Get-Content $outfile  |Out-String

EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -BodyAsHtml -Assunto $subject 

### FINISH ###






