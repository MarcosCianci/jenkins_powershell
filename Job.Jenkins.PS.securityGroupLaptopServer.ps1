################################################################################
# Windows PowerShell 
# 
# Script: Laptop_Server_Group.ps1
# Autor: Marcos Cianci
# Data: 02/07/2013
# Atualização: 	
# 
#
# Versão 1.0: Script verifica se os laptops e servidores são membros do 
#				grupo: BRSP_SERVERS e BRSP_GPO_SCRIPT_GPUPDATE_LAPTOP. 
#				Grupos de permissionamento para GPO e execução de scripts.
#				É verificado se o servidor ou laptop faz parte do grupo, 
#				conforme identificado no descritivo do objeto, se não for
#				o objeto é adcionado como membro do grupo correspondente.
#				Relatório é enviado ao grupo de email suporte@s4bdigital.net.
#
################################################################################

#-------------------- MÓDULOS --------------------#
Import-Module ActiveDirectory

#-------------------- VARIÁVEIS --------------------#
$objFilter = "(objectClass=Computer)"
$objSearch = New-Object System.DirectoryServices.DirectorySearcher
$objSearch.PageSize = 15000
$objSearch.Filter = $objFilter
$objSearch.SearchRoot = "LDAP://dc=SBD,dc=CORP"
$allObj = $objSearch.FindAll()
$date = Get-Date
$group_laptop = "BRSP_GPO_SCRIPT_GPUPDATE_LAPTOP"
$group_server = "BRSP_SERVERS"
$outfile = "E:\usr\util\scripts\logs\Rel_Laptop_Servers_Groups.html"
$img =  "\\sb-dc01\img\s4bdigital.jpg"


#-------------------- CREATE TABLES - REPORT --------------------#
$tabName= "SampleTable"
$table = New-Object system.Data.DataTable "$tabName"
$col1 = New-Object system.Data.DataColumn Hostname,([string]) 
$col2 = New-Object system.Data.DataColumn Type,([string]) 
$col3 = New-Object system.Data.DataColumn Report,([string])
$col4 = New-Object system.Data.DataColumn Group,([string])
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)

#-------------------- CSS --------------------#
$Css='<style>table{margin:auto; width:98%}

     Body{background-color:white; Text-align:Center;}
     th{background-color:black; color:white;}
     td{background-color:Grey; color:white; Text-align:Center;} 
     
     </style>' 


#--------------------  --------------------#

foreach ($Obj in $allObj)
{
	$objItemT = $Obj.Properties
	$CName = $ObjItemT.name
	$comp = $CName
	$description = $objItemT.description
	
	switch -wildcard ($description)
	{
		
	 		"*Laptop*"	{
					#-------------------- Laptop --------------------#	
					$returnL = Get-ADGroupMember -Identity $group_laptop|Select-String "$comp" 
					$d_name_laptop = Get-ADComputer -LDAPFilter "(name=$comp)" |select distinguishedName
					
					$row=$table.NewRow()
					$row.Hostname = "$comp"
					$row.Type = "$description"
					$row.Group = "$group_laptop"
					
					if ($returnL) 
						{
							$row.Report = "Member"
						}
					else 
						{
							$Error.Clear()
							Add-ADGroupMember -Identity $group_laptop -Members $d_name_laptop.distinguishedName
								if($Error)
									{
										$row.Report = "$Error"
									}
								else
									{
										$row.Report = "Add Member"
									}
						}
					$table.Rows.Add($row)	
			}		
	
	
			"*Servidor*" {
	
					#-------------------- Servers --------------------#
							$returnS = Get-ADGroupMember -Identity $group_server|Select-String "$comp" 
							$d_name_server = Get-ADComputer -LDAPFilter "(name=$comp)" |select distinguishedName
							
							$row=$table.NewRow()
							$row.Hostname = "$comp"
							$row.Type = "$description"
							$row.Group = "$group_server"
							
					if ($returnS) 
						{
							$row.Report = "Member"
							
						}
					else 
						{
							$Error.Clear()
							Add-ADGroupMember -Identity $group_server -Members $d_name_server.distinguishedName
							if($Error)
								{
									$row.Report = "$Error"
								}
							else
								{
									$row.Report = "Add Member"
								}
							
						}
					$table.Rows.Add($row)
					
			}
	}
	
}


$log = $table |Select-Object Hostname,Type,Group,Report | Sort-Object Hostname |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Relatório Grupos Servidore - Laptops</h>" | Out-String
$report = ConvertTo-Html -CSSUri $Css -Title "Laptop Servers Group - SBD.CORP " -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log" | Out-String
$report | Out-File $outfile | Out-String


