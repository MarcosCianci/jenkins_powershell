###############################################################
# Windows PowerShell 
# 
# Script: 8.ps1
# Autor: Marcos Cianci
# Data: 03/03/2017
# 
#
# Atualização: 
#  
#
# Versão 1.0: 
#
#
##############################################################

### Modulos ###
Import-Module activedirectory

# Active Directory - Hosts #
$objFilter = "(objectClass=Computer)"
$objSearch = New-Object System.DirectoryServices.DirectorySearcher
$objSearch.PageSize = 15000
$objSearch.Filter = $objFilter
$objSearch.SearchRoot = "LDAP://dc=SBD,dc=CORP"
$allObj = $objSearch.FindAll()

$outfile = "e:\usr\util\scripts\logs\Rel_Audit_Computers_Full.html"
$img = "\\sb-dc01\img\s4bdigital.jpg"
$date = Get-Date -format D
$css = "e:\usr\util\scripts\HtmlReports.css"

### FUNÇÕES ###

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
$col3 = New-Object system.Data.DataColumn Description,([string]) 
$col4 = New-Object system.Data.DataColumn SerialNumber,([string])
$col5 = New-Object system.Data.DataColumn Status,([string])
$col6 = New-Object system.Data.DataColumn Office,([string])
$col7 = New-Object system.Data.DataColumn Processor,([string])
$col8 = New-Object system.Data.DataColumn Memory,([string])
$col9 = New-Object system.Data.DataColumn HardDisk,([string])


$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)
$table.columns.add($col7)
$table.columns.add($col8)
$table.columns.add($col9)


### Starting ###

foreach ($Obj in $allObj)
{
  $objItemT = $Obj.Properties
  $CName = $ObjItemT.name
  $COS = $ObjItemT.operatingsystem
  $comp = $CName
  $Desc = $objItemT.description
   
 $row=$table.NewRow()
 $row.Hostname = "$comp"
 $row.SO = "$COS"
 $row.Description = "$Desc"

 if ( Test-Connection -cn $comp -Count 1 -ErrorAction SilentlyContinue )
   {
                ### Service Tag ###
                $wmic_servicetag = gwmi win32_bios -ComputerName $comp
                $servicetag = $wmic_servicetag.SerialNumber
                $row.SerialNumber = "$servicetag"


                ### CPU, MEMORY, DRIVE DISK ###

                $wmic_processor = Get-WmiObject  -Class win32_processor -ComputerName $comp
                $processor = $wmic_processor.Name
                $Row.Processor = [string]$processor 

               
                $wmic_memory = gwmi -Class win32_operatingsystem -computername $comp  
                $mem = ($wmic_memory.TotalvisibleMemorySize)/1MB
                $memory = [string] $mem.ToString(".00")
                $row.Memory = "$memory GB"
                
                $wmic_disk = get-WmiObject -Class win32_Volume -ComputerName $comp 
                $drive = $wmic_disk |Where-Object { $_.DriveLetter -match "C:" }
                $disk = ($drive.Capacity)/1GB
                $hardisk = $disk.ToString(".00")   
                $row.HardDisk = "$hardisk GB"


                ### Office ###
                
                $obj = Get-WmiObject -class Win32_Product -ComputerName $comp |Where-Object -filterScript { $_.name -match "Microsoft Office Outlook" }
                $version = $obj.version | %{ $_.Split('.')[0]; } |select -First 1
                                              
                switch -wildcard ($version) {

                                    7 { $officename = "Office 97" }
                                    8 { $officename = "Office 98" }
                                    9 { $officename = "Office 2000" }
                                    10 { $officename = "Office XP" }
                                    11 { $officename = "Office 97" }
                                    12 { $officename = "Office 2003" }
                                    13 { $officename = "Office 2007" }
                                    14 {$officename = "Office 2010" }
                                    15 {$officename = "Office 2013" }
                                    16 {$officename = "Office 2016" }
                                    default {$officename = "Microsoft Office not installed"}   
                 }
                  
                 
                 $row.Office = "$officename"  
                 $row.Status = "UP"   
                 
   }     
                    
   else { $row.Status = "Down"}
     
   $table.Rows.Add($row)

}

$log = $table |Select-Object Hostname,SO,Description,Processor,Memory,HardDisk,SerialNumber,Office,Status | Sort-Object Hostname |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Report Summary Computers Dell</h4>" | Out-String 

$report = ConvertTo-Html -CSSUri $css -Title "Report Summary Computers Dell " -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log" |Out-String
$report | Out-File $outfile | Out-String

$emailTo = "suporte@s4bdigital.net"
$subject = "Report Summary - Computers: SerialNumber Dell"
$body_rel = Get-Content $outfile  |Out-String

EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -BodyAsHtml -Assunto $subject 

## End ###