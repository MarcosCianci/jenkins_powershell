#######################################################################
# Windows PowerShell 
# 
# Script: Job.Jenkins.PS.W32TMRsync.ps1
# 
# Autor: Marcos Cianci
# Data: 02/05/2016
# 
#
# Atualização: 	
#               16/02/2017 - Marcos Cianci
#
# Versão 1.0: Script possui a funcionalidade de verificar o status de 
#             sincronismo com o pool de servidores do serviço NTP na 
#             internet e realizar o sincronismo quando ocorrer 
#             disparidade entre os horarios. Envia um relatorio por email 
#             para suporte@s4bdigital.net quando ocorre disparidade informando    
#             o status do sincronismo.    
#
# Versão 2.0: Integração Jenkins  
#             
#########################################################################

### VARIÁVEIS ###
$outfile = "e:\usr\util\scripts\logs\Rel_Test_w32TMResync.html"
$img = ".\img\s4bdigital.jpg"
$date = Get-Date -Format g
$css = "e:\usr\util\scripts\HtmlReports.css"
$Server = "sb-dc01","sb-dc02"


### TABELAS ###
$tabName= "SampleTable"
$table = New-Object system.Data.DataTable "$tabName"
$col1 = New-Object system.Data.DataColumn Server,([string]) 
$col2 = New-Object system.Data.DataColumn RootDispersion,([string])
$col3 = New-Object system.Data.DataColumn Status,([string]) 
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)

### START ###
foreach ( $allserver in $Server ){

    
    $W32TM = w32tm /query /computer:$allserver /Status
    $row=$table.NewRow()
    $row.Server= "$allserver"
	
    $RootD = $W32TM |Where-Object { $_.Contains("Root Dispersion")}
    $RootDispersion = $RootD -replace "Root Dispersion:"

    $row.RootDispersion = [string]"$RootDispersion"
           

        if ( Test-Connection -cn $allserver -Count 1 -ErrorAction SilentlyContinue ){

                if ( $RootDispersion -ge "1.0"){
                    

                   w32tm /resync /computer:$allserver /force 
                   $row.Status = "Resync Success" 

                     }
                else{   
                    echo ""
                    $row.RootDispersion = [string]"$RootDispersion"
                   $row.Status = "Time Ok"
                   }
       
 
           $table.Rows.Add($row)
    
            } 
   
  
   
}

$log = $table |Select-Object Server,RootDispersion,Status | Sort-Object Server |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Relatório - W32TM Domain Controller</h4>" | Out-String
$report = ConvertTo-Html -CSSUri $css -Title "Domain Controller - W32TM Resync" -head "<img src=$img align=middle/> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3>" -body "$log"  | Out-String
$report | Out-File $outfile | Out-String

### FINISH ###