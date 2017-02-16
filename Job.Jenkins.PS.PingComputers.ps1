#######################################################################
# Windows PowerShell 
# 
# Script: Test_Ping_Computers_V2.ps1
# 
# Autor: Marcos Cianci
# Data: 13/04/2015
# 
#
# Atualização: 	
# 
#
# Versão 1.0: Script substitui a versão V1, ele realiza três teste de 
#              pings ( 1 sequencia de 5 pings por testes) em cada teste
#              ele tira a media, se for maior que 3 ms todos os testes 
#              é inserido na tabela e enviado por email um relatorio
#              com todas as estações de trabalho que estão acima da média.
#
#########################################################################

### Modulos ###

Import-Module ActiveDirectory

### Variavés ###

# Active Directory - Hosts #
$objFilter = "(objectClass=Computer)"
$objSearch = New-Object System.DirectoryServices.DirectorySearcher
$objSearch.PageSize = 15000
$objSearch.Filter = $objFilter
$objSearch.SearchRoot = "LDAP://dc=SBD,dc=CORP"
$allObj = $objSearch.FindAll()


$outfile = "e:\usr\util\scripts\logs\Rel_Test_Ping_Computers_V2.html"
$img = "\\sb-dc01\img\s4bdigital.jpg"
$date = Get-Date
$css = "e:\usr\util\scripts\HtmlReports.css"
$SleepTimeOut = "5"
$offcomputer = $null


#-------------------- FUNÇÕES --------------------#
function EnviaEmailUser ($destino,$mensagem,$assunto) 
{
$email = New-Object System.Net.Mail.MailMessage("infra@s4bdigital.net",$destino)
$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
$email.IsBodyHtml = $True
$email.Subject = $assunto
$email.Body = $mensagem
$smtp.Send($email)
}

#-------------------- CREATE TABLES - REPORT --------------------#
$tabName= "SampleTable"
$table = New-Object system.Data.DataTable "$tabName"

$col1 = New-Object system.Data.DataColumn Hostname,([string]) 
$col2 = New-Object system.Data.DataColumn SO,([string])
$col3 = New-Object system.Data.DataColumn Status,([string]) 
$col4 = New-Object system.Data.DataColumn Data, ([string])
$col5 = New-Object system.Data.DataColumn Type,([String])
$col6 = New-Object system.Data.DataColumn Description, ([string])
$col7 = New-Object system.Data.DataColumn IP, ([string])
$col8 = New-Object system.Data.DataColumn HighPing, ([string])
$col9 = New-Object system.Data.DataColumn OffComputer, ([string])
$col10 = New-Object system.Data.DataColumn NormalPing, ([string])

$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.Columns.add($col6) 
$table.Columns.add($col7)
$table.Columns.add($col8)
$table.Columns.add($col9)
$table.Columns.add($col10)


### Starting ###


$avaliable = $null
$notavaliable = $null

# Test #01 #
foreach ($Obj in $allObj)
    {

        $objItemT = $Obj.Properties
        $CName = $ObjItemT.name
        $objItemT = $Obj.Properties
        $CName = $ObjItemT.name
        $COS = $ObjItemT.operatingsystem
        $computer = $CName
        $Desc = $objItemT.description
              
        $ip = [System.Net.DNS]::GetHostAddresses($computer)
        $ip_net = $ip | %{$_.IPAddressToString}

    if ( Test-Connection -ComputerName $computer -Count 1 -ea silentlycontinue )
    {

        # Media Pings 01 #
        $test1 = ( Test-Connection -cn $computer -Count 5 | Measure-Object -Property ResponseTime -Average ).average        
        $res1 = ($test1 -as [int])
        $res1  
        sleep $SleepTimeOut

        # Media Pings 02 #
        $test2 = ( Test-Connection -cn $computer -Count 5 | Measure-Object -Property ResponseTime -Average ).average        
        $res2 = ($test2 -as [int])
        $res2  
        sleep $SleepTimeOut

        # Media Pings 03 #
        $test3 = ( Test-Connection -cn $computer -Count 5 | Measure-Object -Property ResponseTime -Average ).average        
        $res3 = ($test3 -as [int])
        $res3  
    
        if (($res1 -ge "3") -and ($res2 -ge "3") -and ($res3 -ge "3"))
        {

        [array]$avaliable += $computer

        $label= $true

	    $row=$table.NewRow()
		$row.Hostname = "$computer"
		$row.SO = "$COS"
		$row.Data = "$date" 
		$row.Type = "Computer"
		$row.Description = "$Desc"
        $row.Status = "$res1(ms) - $res2(ms) - $res3(ms)"
        $row.IP = "$ip_net"
        $Error.Clear()
		$table.Rows.Add($row)


        }
    
        else {[array]$notavaliable += $computer}   
    } 

    else { [array]$offcomputer += $computer }
   
      
  }


    $row=$table.NewRow()
    $row.OffComputer = $offcomputer.Count
    $row.highping = $avaliable.Count
    $row.NormalPing = $notavaliable.Count
    $table.Rows.Add($row)


if ($label){

$log = $table |Select-Object Hostname,IP,SO,Description,Type,Data,Status | Sort-Object Hostname |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Relatório Computadores - Ping Time(ms)</h4>" | Out-String
$log1 = $table |Select-Object OffComputer,HighPing,NormalPing | ConvertTo-Html -Fragment -As table -PreContent "<h4>Report Computers Total</h4>" | Out-String

$report = ConvertTo-Html -CSSUri $css -Title "Computers Active Directory - Test Ping Time(ms) Computers" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log $log1" | Out-String
$report | Out-File $outfile | Out-String


$emailTo = "suporte@s4bdigital.net"
$subject = "Relatório Computadores do Active Directory - Ping Time(ms) V2 "
$body_rel = Get-Content $outfile  |Out-String

EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -BodyAsHtml -Assunto $subject 

}
   


### END ###