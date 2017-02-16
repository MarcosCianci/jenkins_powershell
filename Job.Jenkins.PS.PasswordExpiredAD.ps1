###############################################################
# Windows PowerShell 
# Script: job.Jenkins.PS.PasswordExpiredAD.ps1
#
# Autor: Marcos Cianci
# Data: 30/01/2013
#
# Atualização: 06/02/2013 - Marcos Cianci
#              17/07/2014 - Marcos Cianci	
# 
#
# Versão 1.0: Script verifica as senhas dos usuários que estão 
#				para expirar em 5 dias. Será enviado ao usuário 
#				notificação por email e um relatório das contas
#				expiradas para equipe Suporte e Infraestrutura	
#
# Versão 2.0: Implementado a função que realiza a validação do
#				email cadastrado nas propriedades do usuario. 
#				Inserido também a verificação de erros no envio
#				das mensagens para o usuario caso o email esteja
#				cadastrado de forma incorreta, enviado no relatorio
#				para equipe de infraestrutura
# 
# Versão 3.0: Alterado as configurações de geração dos relatorios
#               para o formato em html, assim como o envio dos
#               e-mails para os usuário em hmtl. 
#
# Versão 4.0: Integração Jenkins
#
###################################################################

### CARREGANDO MODULO AD ###
Import-Module ActiveDirectory

### PRAZO EXPIRAÇÃO DE SENHAS USUARIO - 90 DIAS ###
$maxdays=(Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.TotalDays
$day_max="5"
$day_min="0"

### Variavés ###
$outfile = "e:\usr\util\scripts\logs\Rel_Senhas_Expiradas_AD.html"
$img = "\\sb-dc01\img\s4bdigital.jpg"
$date = Get-Date -format D
$css = "e:\usr\util\scripts\HtmlReports.css"

### CREATE TABLES - REPORT ###
$tabName= "SampleTable"
$table = New-Object system.Data.DataTable "$tabName"
$col1 = New-Object system.Data.DataColumn Usuario,([string]) 
$col2 = New-Object system.Data.DataColumn DataExpiração,([string])
$col3 = New-Object system.Data.DataColumn Expira,([string]) 
$col4 = New-Object system.Data.DataColumn Email,([string])
$col5 = New-Object system.Data.DataColumn Status,([String])
$col6= New-Object system.Data.DataColumn ReportEmail,([String])
$table.columns.add($col1)
$table.columns.add($col2)
$table.columns.add($col3)
$table.columns.add($col4)
$table.columns.add($col5)
$table.columns.add($col6)

### FUNÇÃO ENVIO E-MAIL ###
function EnviaEmailUser ($destino,$mensagem,$assunto) 
{
$email = New-Object System.Net.Mail.MailMessage("suporte@s4bdigital.net",$destino)
$email.Subject = $assunto
$email.Body = $mensagem
$email.Priority = 2
$email.IsBodyHtml = $True
$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
$smtp.Send($email)
}

### FUNÇÃO VALIDAÇÃO E-MAIL CADASTRO AD ###
function ValidateEmail {
	param([string]$address)
	($address -as [System.Net.Mail.MailAddress]).address -eq $address -and $address -ne $null
}

### BUSCA USUARIO E ENVIO E-MAIL ###
(Get-ADUser -filter {(Description -notlike "IfYouWantToExclude*") -and (Enabled -eq "True") -and (PasswordNeverExpires -eq "False")} -properties *) | Sort-Object pwdLastSet |
foreach-object {

	$lastset=Get-Date([System.DateTime]::FromFileTimeUtc($_.pwdLastSet))
	$expires=$lastset.AddDays($maxdays).ToShortDateString()
	$daystoexpire=[math]::round((New-TimeSpan -Start $(Get-Date) -End $expires).TotalDays)
	$samname=$_.samaccountname
	$mail=$_.mail
	$validaemail= ValidateEmail ("$mail")
	$firstname=$_.GivenName
	$cn=$_.cn
	
	if (($daystoexpire -le $day_max) -and ($daystoexpire -ge $day_min ))
		{
    		$ThereAreExpiring=$true

			if ($validaemail -eq $true)
			{
			$emailTo= $mail
    		$subject="$firstname, sua senha irá expirar em $daystoexpire dia(s)"
    		$body="

<HEAD>
 <STYLE type=text/css>
  BODY {text-align: left}
  p {font-size: 14px; font-family: Arial} </STYLE>
  
<p><H4> $cn, </H4></p>

<p>Sua senha irá expirar em $daystoexpire dia(s).</p>

<h4><p style=color:red >!!! Por favor, leia com atenção !!!</p></h4>

<p>Esta é a senha do Windows usada para logar-se em sua Estação de trabalho, se você não alterar sua senha antes do prazo de expiração acima, ela será automaticamente bloqueada, impedindo o seu acesso a internet, email e rede.</p>
	
<p>Siga os passos abaixo para alterar sua senha:</p>

	<p><ul>1. Pressione CTRL+ALT+DEL;</ul></p>
	<p><ul>2. Na tela clique em 'Alterar sua senha';</ul></p>
	<p><ul>3. Digite sua senha antiga,e digite duas vezes sua nova senha (aconselhável não repetir as ultimas três senhas utilizadas anteriormente);</ul></p>
	<p><ul>4. Após ter realizado a alteração, surgirá a mensagem na tela informando que sua senha foi alterada com sucesso;</ul></p>

<p>Dúvidas ou dificuldades para alterar sua senha, entre em contato com Suporte Técnico pelo ramal 8010 ou acesse http://intranet.localdomain/ocomon/ para abrir um ticket.</p> 

<p>Atenciosamente,</p>
<h4><p>Equipe Service Desk - S4BDigital</p></h4>
<img src=$img alt='Image Error' vspace='10x' >

 <h4><p>*** Este e-mail é gerado automaticamente - Por favor, não responda ***</p></h4>

</BODY>
"
	EnviaEmailUser -Destino $emailTo -Mensagem $body -Assunto $subject 
	
	$row=$table.NewRow()
	$row.Usuario = "$cn" 
	$row.DataExpiração = "$expires"
	$row.Expira = "$daystoexpire(dias)" 
	$row.Email = "$mail"
	
				if ($Error -ceq $Error)
					{			
					$row.Status = "Envio Email Error"
					$row.ReportEmail="$Error"
					$table.Rows.Add($row)
					}
	
				else
					{
					$row.Status = "Envio Email Ok"
					$table.Rows.Add($row)
					}
			$Error.Clear()
			
			}
			else
			{
				$row=$table.NewRow()
				$row.Usuario = "$cn" 
				$row.DataExpiração = "$expires"
				$row.Expira = "$daystoexpire(dias)" 
				$row.Email = "$mail"

				if ($mail -eq $null)
					{
					$row.Status = "Email não cadastrado no AD"
					$table.Rows.Add($row)
					}
				else{
					$row.Status = "Formato Email Incorreto"
					$table.Rows.Add($row)
				}
		
			}
		}
}

$log = $table |Select-Object Usuario,DataExpiração,Expira,Email,Status |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Relatório Usuários - Senhas Expiradas</h4>" | Out-String 
$log1 = $table |Select-Object ReportEmail |ConvertTo-Html -Fragment -As Table -PreContent "<h4>Report Email</h4>" | Out-String

$report = ConvertTo-Html -CSSUri $css -Title "Active Directory - Senhas Expiradas" -head "<img src=$img align=middle> <H2>Depart. InfraEstrutura e Suporte</H2> <h3>Data:$date</h3> Servidor: SB-DC01" -body "$log $log1" |Out-String
$report | Out-File $outfile | Out-String

### ENVIO RELÁTORIO POR EMAIL - EQUIPE INFRAESTRUTURA E SUPORTE ###

if ($ThereAreExpiring) {
	
	$emailTo = "suporte@s4bdigital.net"
	$subject = "Relatório Usuários - Senhas Expiradas"
	$body_rel = get-content $outfile | out-string
    
	EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -BodyAsHtml -Assunto $subject 
}
else {
	$emailTo = "suporte@s4bdigital.net"
	$subject = "Relatório Usuários - Senhas Expiradas"
	$body_rel = "Não há usuários com senhas expiradas com menos de $day_max(dias)"
	EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -BodyAsHtml -Assunto $subject 
}