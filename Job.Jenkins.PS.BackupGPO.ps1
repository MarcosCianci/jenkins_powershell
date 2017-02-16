#################################################################
# Windows PowerShell 
# 
# Script: Backup_GPO_Server.ps1
# Autor: Marcos Cianci
# Data: 05/03/2013
# 
# Atualização: 	
#               16/02/2017 - Marcos Cianci    
#
# Versão 1.0: Script realiza o backup das Politicas de Grupos 
#				configuradas no Dominio SBD.CORP para usuários,
#				computadores, grupos e unidades organizacionais.
#
# Versão 2.0: Script integração com Jenkins
#
##################################################################

### CARREGANDO MODULOS ###
Import-Module grouppolicy

### VARIÁVEL ###
$bkp_dir ="H:\bkp_GPO"
$date = Get-Date
$outfile = "E:\usr\util\scripts\logs\Rel_Backup_GPO.html"
$img =  "\\sb-dc01\img\s4bdigital.jpg"
$css = "E:\usr\util\scripts\HtmlReports.css"
$date_del = (Get-Date) - (New-TimeSpan -day 5)


### STARTING ###

$bkp_gpo = Get-GPO -all | Backup-GPO -path $bkp_dir |ConvertTo-Html -Fragment -As Table -PreContent "<h2>Relatório</h2>" | Out-String
$report = ConvertTo-Html -CSSUri $css -Title "GPO´s Backup - Active Direcotry" -head "<img src=\\sb-dc01\img\s4bdigital.jpg align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$bkp_gpo" 
$report | Out-File $outfile | Out-String

### FINISH ###

