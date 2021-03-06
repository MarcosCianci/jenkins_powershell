###############################################################
# Windows PowerShell 
# 
# Script: Query_EventLogs.ps1
# Autor: Marcos Cianci
# Data: 04/03/2012
# 
#
# Atualização: 
#
#			Marcos Cianci - 11/09/2014
#           Marcos Cianci - 29/04/2015  
# 
# 
#
# Versão 1.0: 
#
#			Script verifica os logs "Security" referente as alterações
#			realizadas no Active Direcotry no servidor
#			Controlador de Dominio - S4B-DC01
#
#
#			* Evento 5141 - Log de Eventos de Objetos deletados no Active
#			Directory. (Usuários,Grupos,Computadores, Unidades Organizacionais
#			e Politicas de Grupos)
#
#			* Evento 5137 - Log de Eventos de Objetos criados no Active
#			Directory. (Usuários,Grupos,Computadores, Unidades Organizacionais
#			e Politicas de Grupos)
#
#			* Evento 5139 - Log de Eventos de Objetos movidos no Active
#			Directory. (Usuários,Grupos,Computadores, Unidades Organizacionais
#			e Politicas de Grupos) 
#
#			* Evento 5136 - Log de Eventos de Objetos Modificados no Active
#			Directory. (Usuários,Grupos,Computadores, Unidades Organizacionais
#			e Politicas de Grupos)
#
#			Envia por e-mail para suporte@s4bdigital.net as alterações realizadas,
#			informando o evento, o horario gerado, o usuario, objeto e entre outras 
#			informações no periodo de  horas atrás. O agendamento será configurado 
#			no Task Schedule para ser executado 4 vez por dia a cada 6 horas.
#
#			Verificar nos relatorios a referencia de objetos modificados
#
# Versão 1.1: 
#
#			Foi inserido no script a funcionalidade de verificar objetos criados no
#			container Computer - "CN=Computers,DC=SBD,DC=CORP" que não foram 
#			atribuidos para OU correta. É enviado um email para equipe de Suporte e
#			Infraestrutura informando que existe um equipamento não atrelado a OU de 
#			departamento correspondente.
#
########################################################################################

### Modulos ###
Import-Module activedirectory


### CREATE TABLES ### 

## DELETE ##
$table = New-Object system.Data.DataTable "TableDeleted"
$coldel1 = New-Object system.Data.DataColumn AlertName ,([string])
$coldel2 = New-Object system.Data.DataColumn Message ,([string])
$coldel3 = New-Object system.Data.DataColumn WhoChanges ,([string])
$coldel4 = New-Object system.Data.DataColumn ObjectDeleted ,([string])
$coldel5 = New-Object system.Data.DataColumn ObjectPath ,([string])
$coldel6 = New-Object system.Data.DataColumn ObjectType ,([string])
$coldel7 = New-Object system.Data.DataColumn TimeCreated ,([string])
$coldel8 = New-Object system.Data.DataColumn Server ,([string])
$coldel9 = New-Object system.Data.DataColumn CorrelationID ,([string])
$table.columns.add($coldel1)
$table.columns.add($coldel2)
$table.columns.add($coldel3)
$table.columns.add($coldel4)
$table.columns.add($coldel5)
$table.columns.add($coldel6)
$table.columns.add($coldel7)
$table.columns.add($coldel8)
$table.columns.add($coldel9)


## CREATED ##
$table2 = New-Object system.Data.DataTable "TableCreated"
$colcrd1 = New-Object system.Data.DataColumn AlertName ,([string])
$colcrd2 = New-Object system.Data.DataColumn Message ,([string])
$colcrd3 = New-Object system.Data.DataColumn WhoChanges ,([string])
$colcrd4 = New-Object system.Data.DataColumn ObjectNew ,([string])
$colcrd5 = New-Object system.Data.DataColumn ObjectPath ,([string])
$colcrd6 = New-Object system.Data.DataColumn ObjectType ,([string])
$colcrd7 = New-Object system.Data.DataColumn TimeCreated ,([string])
$colcrd8 = New-Object system.Data.DataColumn Server ,([string])
$colcrd9 = New-Object system.Data.DataColumn CorrelationId ,([string])
$table2.columns.add($colcrd1)
$table2.columns.add($colcrd2)
$table2.columns.add($colcrd3)
$table2.columns.add($colcrd4)
$table2.columns.add($colcrd5)
$table2.columns.add($colcrd6)
$table2.columns.add($colcrd7)
$table2.columns.add($colcrd8)
$table2.columns.add($colcrd9)

## MOVED ##
$table3 = New-Object system.Data.DataTable "TableMoved"
$colmov1 = New-Object system.Data.DataColumn AlertName ,([string])
$colmov2 = New-Object system.Data.DataColumn Message ,([string])
$colmov3 = New-Object system.Data.DataColumn WhoChanges ,([string])
$colmov4 = New-Object system.Data.DataColumn ObjPathNew ,([string])
$colmov5 = New-Object system.Data.DataColumn ObjPathOld ,([string])
$colmov6 = New-Object system.Data.DataColumn Object ,([string])
$colmov7 = New-Object system.Data.DataColumn ObjectType ,([string])
$colmov8 = New-Object system.Data.DataColumn TimeCreated ,([string])
$colmov9 = New-Object system.Data.DataColumn Server ,([string])
$colmov10 = New-Object system.Data.DataColumn CorrelationId ,([string])
$table3.columns.add($colmov1)
$table3.columns.add($colmov2)
$table3.columns.add($colmov3)
$table3.columns.add($colmov4)
$table3.columns.add($colmov5)
$table3.columns.add($colmov6)
$table3.columns.add($colmov7)
$table3.columns.add($colmov8)
$table3.columns.add($colmov9)
$table3.columns.add($colmov10)

## MODIFIED ##

$table4 = New-Object system.Data.DataTable "TableMod"
$colmod1 = New-Object system.Data.DataColumn AlertName ,([string])
$colmod2 = New-Object system.Data.DataColumn Message ,([string])
$colmod3 = New-Object system.Data.DataColumn WhoChanges ,([string])
$colmod4 = New-Object system.Data.DataColumn LDAPName ,([string])
$colmod5 = New-Object system.Data.DataColumn Value ,([string])
$colmod6 = New-Object system.Data.DataColumn Type ,([string])
$colmod7 = New-Object system.Data.DataColumn ObjectPath ,([string])
$colmod8 = New-Object system.Data.DataColumn ObjectType ,([string])
$colmod9 = New-Object system.Data.DataColumn TimeCreated ,([string])
$colmod10 = New-Object system.Data.DataColumn Server ,([string])
$colmod11 = New-Object system.Data.DataColumn Object ,([string])
$colmod12 = New-Object system.Data.DataColumn CorrelationID ,([string])
$table4.columns.add($colmod1)
$table4.columns.add($colmod2)
$table4.columns.add($colmod3)
$table4.columns.add($colmod4)
$table4.columns.add($colmod5)
$table4.columns.add($colmod6)
$table4.columns.add($colmod7)
$table4.columns.add($colmod8)
$table4.columns.add($colmod9)
$table4.columns.add($colmod10)
$table4.columns.add($colmod11)
$table4.columns.add($colmod12)


## COMPUTER CONTAINER OU COMPUTERS ##

$table5 = New-Object system.Data.DataTable "TableComputer"
$colcomp1 = New-Object system.Data.DataColumn Name ,([string])
$colcomp2 = New-Object system.Data.DataColumn ObjectClass ,([string])
$colcomp3 = New-Object system.Data.DataColumn DistinguishedName ,([string])
$table5.columns.add($colcomp1)
$table5.columns.add($colcomp2)
$table5.columns.add($colcomp3)

### VARIÁVEIS GLOBAIS ###

$hours = "6"
$date = (Get-Date) - (New-TimeSpan -Hours $hours)
$emailTo = "suporte@s4bdigital.net"
$Server = "SB-DC01"
$Subject="Alerta - Relatório de Objetos no Active Directory - SBD.CORP"
$img = "E:\img\s4bdigital.jpg"
$css = "e:\usr\util\scripts\HtmlReports.css"

### FUNÇÕES ###

# ENVIO E-MAIL SUPORTE 

function EnviaEmailUser ($destino,$mensagem,$assunto){
		$email = New-Object System.Net.Mail.MailMessage("infra@s4bdigital.net",$destino)
		$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
        $email.Subject = $assunto
		$email.Body = $mensagem
		$email.Priority = 2
        $email.IsBodyHtml = $True
		$smtp.Send($email)
}

### PROCESSANDO ... ###


#########################################################################
## EVENTO 5141 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) DELETADO NO AD ##
#########################################################################

	$EventLog = Get-WinEvent -FilterHashtable @{LogName="security";id="5141"}
	$EventLog |Where-Object {($_.Message -match "deleted")-and ($_.TimeCreated -ge $date)} |ForEach-Object {

	$class = [string](select-string -inputobject $_.message -allmatches -pattern 'Class:\s+(.+)').matches[0].groups[1].value;
	$dn = [string](select-string -inputobject $_.message -allmatches -pattern 'DN:\s+(.+)').matches[0].groups[1].value;
	$AccountName = [string](select-string -inputobject $_.message -allmatches -pattern 'Account Name:\s+(.+)').matches[0].groups[1].value;
	$correlationId = [string](select-string -inputobject $_.message -allmatches -pattern 'Correlation ID:\s+(.+)').matches[0].groups[1].value;
	$TimeCreated = [string]$_.TimeCreated;
	$message = "A directory service object was deleted."
	$cn = $dn |Select-String -pattern "CN=([A-Za-z \.-_]+)," | foreach {$_.matches} |foreach {$_.groups[1].value}|Select-Object -Unique 
	$AlertName="Objeto deletado no Active Directory - SBD.CORP" 
	
	
	switch -regex ($class.ToLower())
	{
		"[user]"{
		$label_del="$true"
		$row = $table.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectDeleted = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table.Rows.Add($row)
		break;
		}
		
		"[group]"{
		$label_del="$true";
		$row = $table.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectDeleted = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table.Rows.Add($row)
		break;
		}
		
		"[computer]"{
		$label_del="$true";
		$row = $table.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectDeleted = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table.Rows.Add($row)
		break;
		}
		
		"[organizationunit]"{
		$label_del="$true";
		$row = $table.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectDeleted = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table.Rows.Add($row)
		break;
		}
		"[groupPolicyContainer]"{
		$label_del="$true";
		$row = $table.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectDeleted = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table.Rows.Add($row)
		break;
		}
		
		
	}
} 
## ENVIO EMAIL - REPORT DELETE ##

		if($label_del)
		{
		#$body = $table |Select-Object AlertName,Message,WhoChanges,ObjectDeleted,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |Format-list -Property * |Out-String -Width 4096
        				
		$outfiledel = "e:\usr\util\scripts\logs\Rel_Eventos_Deletados_AD.html"
        $body = $table |Select-Object AlertName,Message,WhoChanges,ObjectDeleted,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |ConvertTo-Html -Fragment -As List -PreContent "<h4>EVENTO 5141 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) DELETADO NO AD</h4>" | Out-String 

        $report = ConvertTo-Html -CSSUri $css -Title "EVENTO 5141 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) DELETADO NO AD" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$body" |Out-String
        $report | Out-File $outfiledel | Out-String
        
      
        EnviaEmailUser -Destino $emailTo -Mensagem $report -Assunto "Alerta - Relatório Objetos Deletados no Active Directory - SBD.CORP"
        
		}
        
		else {echo ""}
	
########################################################################
## EVENTO 5137 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) CRIADO NO AD ###
########################################################################
	
	$EventLog= Get-WinEvent -FilterHashtable @{LogName="security";id="5137"}
	
	$EventLog |Where-Object {($_.Message -match "created")-and ($_.TimeCreated -ge $date)} |ForEach-Object {
	
	$class = [string](select-string -inputobject $_.message -allmatches -pattern 'Class:\s+(.+)').matches[0].groups[1].value;
	$dn = [string](select-string -inputobject $_.message -allmatches -pattern 'DN:\s+(.+)').matches[0].groups[1].value;
	$AccountName = [string](select-string -inputobject $_.message -allmatches -pattern 'Account Name:\s+(.+)').matches[0].groups[1].value; 
	$correlationId = [string](select-string -inputobject $_.message -allmatches -pattern 'Correlation ID:\s+(.+)').matches[0].groups[1].value;
	$TimeCreated = [string]$_.TimeCreated;
	$message = "A directory service object was created."
	$cn = $dn |Select-String -pattern "CN=([A-Za-z \.-_]+)," | foreach {$_.matches} |foreach {$_.groups[1].value}|Select-Object -Unique
	$AlertName = "Objeto Criado no Active Directory - SBD.CORP"
	
	switch -regex ($class.ToLower())
	{
		"[user]"{
		$label_created="$true"
		$row = $table2.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectNew = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table2.Rows.Add($row)
		break;
		}
		
		"[group]"{
		$label_created="$true";
		$row = $table2.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectNew = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table2.Rows.Add($row)
		break;
		}
		
		"[computer]"{
		$label_created="$true";
		$row = $table2.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectNew = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table2.Rows.Add($row)
		break;
		}
		
		"[organizationunit]"{
		$label_created="$true";
		$row = $table2.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectNew = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table2.Rows.Add($row)
		break;
		}
		"[groupPolicyContainer]"{
		$label_created="$true";
		$row = $table2.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjectNew = "$cn";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table2.Rows.Add($row)
		break;
		}
		
	}
}
	
## ENVIO EMAIL - REPORT CREATED ## 

		if($label_created)
		{
		
        #$body = $table2 |Select-Object AlertName,Message,WhoChanges,ObjectNew,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |Format-list -Property * |Out-String -Width 4096				
		        
        $outfilecreate = "e:\usr\util\scripts\logs\Rel_Eventos_Criados_AD.html"
        $body = $table2 |Select-Object AlertName,Message,WhoChanges,ObjectDeleted,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |ConvertTo-Html -Fragment -As List -PreContent "<h4>EVENTO 5137 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) CRIADO NO AD</h4>" | Out-String 

        $report = ConvertTo-Html -CSSUri $css -Title "EVENTO 5137 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) CRIADO NO AD" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$body" |Out-String
        $report | Out-File $outfilecreate | Out-String
      
        EnviaEmailUser -Destino $emailTo -Mensagem $report -Assunto "Alerta - Relatório Objetos Criados no Active Directory - SBD.CORP"
		
        }
		else {echo ""}

########################################################################
## EVENTO 5139 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MOVIDOS NO AD ##
########################################################################
	
	$EventLog= Get-WinEvent -FilterHashtable @{LogName="security";id="5139"}
	
	$EventLog |Where-Object {($_.Message -match "moved")-and ($_.TimeCreated -ge $date)} |ForEach-Object {
	
	$class = [string](select-string -inputobject $_.message -allmatches -pattern 'Class:\s+(.+)').matches[0].groups[1].value;
	$dn = [string](select-string -inputobject $_.message -allmatches -pattern 'DN:\s+(.+)').matches[0].groups[1].value;
	$AccountName = [string](select-string -inputobject $_.message -allmatches -pattern 'Account Name:\s+(.+)').matches[0].groups[1].value; 
	$TimeCreated = [string]$_.TimeCreated;
	$message = "A directory service object was moved."
	$cn = $dn |Select-String -pattern "CN=([A-Za-z \.-_]+)," | foreach {$_.matches} |foreach {$_.groups[1].value}|Select-Object -Unique
	$dnold = [string](select-string -inputobject $_.message -allmatches -pattern 'Old DN:\s+(.+)').matches[0].groups[1].value;
	$dnnew = [string](select-string -inputobject $_.message -allmatches -pattern 'New DN:\s+(.+)').matches[0].groups[1].value;
	$correlationId = [string](select-string -inputobject $_.message -allmatches -pattern 'Correlation ID:\s+(.+)').matches[0].groups[1].value;
	$AlertName = "Objeto Movido no Active Directory - SBD.CORP"
	
	switch -regex ($class.ToLower())
	{
		"[user]"{
		$label_mov="$true"
		$number="$cont";
		$row = $table3.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjPathOld = "$dnold";
		$row.ObjPathNew = "$dnnew";
		$row.Object = "$cn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table3.Rows.Add($row)
		break;
		}
		
		"[group]"{
		$label_mov="$true";
		$row = $table3.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjPathOld = "$dnold";
		$row.ObjPathNew = "$dnnew";
		$row.Object = "$cn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table3.Rows.Add($row)
		break;
		}
		
		"[computer]"{
		$label_mov="$true";
		$row = $table3.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjPathOld = "$dnold";
		$row.ObjPathNew = "$dnnew";
		$row.Object = "$cn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table3.Rows.Add($row)
		break;
		}
		
		"[organizationunit]"{
		$label_mov="$true";
		$row = $table3.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjPathOld = "$dnold";
		$row.ObjPathNew = "$dnnew";
		$row.Object = "$cn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table3.Rows.Add($row)
		break;
		}
		"[groupPolicyContainer]"{
		$label_mov="$true";
		$row = $table3.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.ObjPathOld = "$dnold";
		$row.ObjPathNew = "$dnnew";
		$row.Object = "$cn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table3.Rows.Add($row)
		break;
		}
		
		
		
	}
}
## ENVIO EMAIL - REPORT MOVED ## 

		if($label_mov)
		{
		#$body = $table3 |Select-Object AlertName,Message,WhoChanges,ObjPathOld,ObjPathNew,Object,ObjectType,TimeCreated,Server,CorrelationId |Format-list -Property * |Out-String -Width 4096				
		
        $outfilemov = "e:\usr\util\scripts\logs\Rel_Eventos_Movidos_AD.html"
        $body = $table3 |Select-Object AlertName,Message,WhoChanges,ObjPathOld,ObjPathNew,ObjectDeleted,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |ConvertTo-Html -Fragment -As List -PreContent "<h4>EVENTO 5139 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MOVIDOS NO AD </h4>" | Out-String 

        $report = ConvertTo-Html -CSSUri $css -Title "EVENTO 5139 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MOVIDOS NO AD " -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$body" |Out-String
        $report | Out-File $outfilemov | Out-String
        
        EnviaEmailUser -Destino $emailTo -Mensagem $report -Assunto "Alerta - Relatório Objetos Movidos no Active Directory - SBD.CORP"
		}
        
		else {echo ""}
		
		
############################################################################
## EVENTO 5136 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MODIFICADOS NO AD ##
############################################################################


	$EventLog= Get-WinEvent -FilterHashtable @{LogName="security";id="5136"} 
	$cont="0";
	
	$EventLog |Where-Object {($_.Message -match "modified")-and ($_.TimeCreated -ge $date)} |ForEach-Object {
	
	$class = [string](select-string -inputobject $_.message -allmatches -pattern 'Class:\s+(.+)').matches[0].groups[1].value;
	$dn = [string](select-string -inputobject $_.message -allmatches -pattern 'DN:\s+(.+)').matches[0].groups[1].value;
	$AccountName = [string](select-string -inputobject $_.message -allmatches -pattern 'Account Name:\s+(.+)').matches[0].groups[1].value; 
	$TimeCreated = [string]$_.TimeCreated;
	$message = "A directory service object was modified."
	$LDAPName = [string](select-string -inputobject $_.message -allmatches -pattern 'LDAP Display Name:\s+(.+)').matches[0].groups[1].value;
	$value = [string](select-string -inputobject $_.message -allmatches -pattern 'Value:\s+(.+)').matches[0].groups[1].value;
	$type = [string](select-string -inputobject $_.message -allmatches -pattern 'Type:\s+(.+)').matches[0].groups[1].value;
	$correlationId = [string](select-string -inputobject $_.message -allmatches -pattern 'Correlation ID:\s+(.+)').matches[0].groups[1].value;	
	$message = "A directory service object was modified."
	$AlertName = "Objeto Modificados no Active Directory - SBD.CORP"
	
	switch -regex ($class.ToLower())
	{
		"[user]"{
		$label_mod="$true"
		$row = $table4.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.LDAPName = "$LDAPName";
		$row.value = "$value"; 
		$row.type = "$type";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table4.Rows.Add($row)
		break;
		}
		
		"[group]"{
		$label_mod="$true";
		$row = $table4.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.LDAPName = "$LDAPName";
		$row.value = "$value"; 
		$row.type = "$type";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table4.Rows.Add($row)
		break;
		}
		
		"[computer]"{
		$label_mod="$true";
		$row = $table4.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.LDAPName = "$LDAPName";
		$row.value = "$value"; 
		$row.type = "$type";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table4.Rows.Add($row)
		break;
		}
		
		"[organizationunit]"{
		$label_mod="$true";
		$row = $table4.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.LDAPName = "$LDAPName";
		$row.value = "$value"; 
		$row.type = "$type";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table4.Rows.Add($row)
		break;
		}
		"[groupPolicyContainer]"{
		$label_mod="$true";
		$row = $table4.NewRow();
		$row.AlertName = "$AlertName";
		$row.Message = "$message";
		$row.WhoChanges = "$AccountName";
		$row.LDAPName = "$LDAPName";
		$row.value = "$value"; 
		$row.type = "$type";
		$row.ObjectPath = "$dn";
		$row.ObjectType = "$class";
		$row.TimeCreated = "$TimeCreated";
		$row.Server = "$Server";
		$row.CorrelationId = "$CorrelationId";
		$table4.Rows.Add($row)
		break;
		}
	}
} 

## ENVIO EMAIL - REPORT MODIFIED ## 

		if($label_mod)
		{
		#$body = $table4 |Select-Object AlertName,Message,WhoChanges,LDAPName,Value,Type,ObjectPath,ObjectType,TimeCreated,Server,CorrelationID |Format-list -Property * |Out-String -Width 4096  				
		
        $outfilemod = "e:\usr\util\scripts\logs\Rel_Eventos_Modificados_AD.html"
        $body = $table4 |Select-Object AlertName,Message,WhoChanges,ObjectDeleted,ObjectPath,ObjectType,TimeCreated,Server,CorrelationId |ConvertTo-Html -Fragment -As List -PreContent "<h4>EVENTO 5136 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MODIFICADOS NO AD </h4>" | Out-String 

        $report = ConvertTo-Html -CSSUri $css -Title "EVENTO 5136 - OBJETO (USUÁRIO,COMPUTADOR,GRUPO E OU) MODIFICADOS NO AD" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$body" |Out-String
        $report | Out-File $outfilemod | Out-String
          
        EnviaEmailUser -Destino $emailTo -Mensagem $report -Assunto "Alerta - Relatório Objetos Modificados no Active Directory - SBD.CORP" 
		}
        
		else {echo ""}
		
##################################################################
### REPORT COMPUTERS CONTAINER - "CN=Computers,DC=SBD,DC=CORP" ###
##################################################################

		$rel_info = Get-ADComputer -filter {(ObjectClass -eq "computer")} -SearchBase "CN=Computers,DC=SBD,DC=CORP"
		
		if ($rel_info) { $label_rel = $true; }

		foreach ( $rel in $rel_info ) {
			
			$row = $table5.NewRow();
			$row.Name = $rel.Name;
			$row.ObjectClass = $rel.ObjectClass;
			$row.DistinguishedName = $rel.DistinguishedName;
			$table5.Rows.Add($row);
			
		}

## ENVIO E-MAIL - REPORT ###

		if($label_rel)
		{
		$outfilecomp = "e:\usr\util\scripts\logs\Rel_Container_Computers.html"
		$emailTo1="suporte@s4bdigital.net"
        $body = $table5 |Select-Object Name,ObjectClass,DistinguishedName |ConvertTo-Html -Fragment -As Table -PreContent "<h4> <span style=color:'#DE1D1D'> Alerta!!! Favor verificar há computadores no Container:Computers!!!</h4>" | Out-String 
        $report = ConvertTo-Html -CSSUri $css -Title "Report Computers - Container: Computers" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: S4B-DC01" -body "$body" |Out-String
        $report | Out-File $outfilecomp | Out-String
        EnviaEmailUser -Destino $emailTo1 -Mensagem $report -Assunto "Alerta - Computador no Container: Computers - SBD.CORP" 
		}
		else
		{ "" }




