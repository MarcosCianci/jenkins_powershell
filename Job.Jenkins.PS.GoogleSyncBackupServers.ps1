#######################################################################
# Windows PowerShell 
# 
# Script: Job.Jenkins.PS.GoogleSyncBackupServers.ps1
#
# Autor: Marcos Cianci
# Data: 245/04/2017
# 
# Atualização: 	
# 
#
# Versão 1.0: 
#
#
#
#########################################################################


### VARIÁVEIS ###

$outfile = "D:\usr\util\scripts\Logs\Rel_Google_Sync_Backup_Servers.html"
$img = "\\sb-dc01\img\s4bdigital.jpg"
$css = "D:\usr\util\scripts\HtmlReports.css"

$date = Get-Date -Format D
$DATA = Get-Date -UFormat "%d%m%Y" 
$datetobkp = (get-date).addDays($Daysbkp)
$Daysbkp = "-3"

$server = "S4B-ACESSO"

$BKPWINSERVER = "E:\WinBKP" 
$BKPGDRIVE = "E:\BkpGDrive"

$Daysback = "-1"
$CurrentDate = Get-Date
$timestamp = Get-Date -Format ddMMyyyy_HHmmss
$DatetoDelete = $CurrentDate.AddDays($Daysback)

$securepassword = ConvertTo-SecureString -String $env:SrvPassword -AsPlainText -Force 
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $env:SrvUser, $securepassword 



### FUNÇÕES ### 
## E-MAIL ##

function EnviaEmailUser ($destino,$mensagem,$assunto) {

		$email = New-Object System.Net.Mail.MailMessage("suporte@s4bdigital.net",$destino)
		$smtp = New-Object Net.Mail.SmtpClient("mail.s4bdigital.net","25")
		$email.Subject = $assunto
		$email.Body = $mensagem
		$email.Priority = 2
		$email.IsBodyHtml = $True
		$smtp.Send($email)
}


###### START ######



foreach ($allserver in $server) {
  
### COMPACTANDO BACKUP ###
       
    $ARQ = Invoke-Command -ComputerName $allserver -ScriptBlock { Get-ChildItem $BKPWINSERVER } -Credential $cred   
    $DIRZIP = [string]"$server"
    $DATAZIP = $DATA   
    $ARQZIP = $DIRZIP + '_' + $DATAZIP + '.' + 'zip' 	 
 
    Invoke-Command -ComputerName $allserver -ScriptBlock { CMD.EXE /C "C:\Program Files\7-Zip\7z.exe" a "$BKPGDRIVE\$ARQZIP" "$BKPWINDOWS\$ARQ"  } -Credential $cred
    $FILEZIP = Invoke-Command -ComputerName $allserver -ScriptBlock { (Get-ChildItem -Filter *.zip $BKPGDRIVE |Where-Object {( $_.LastWriteTime -ge $CurrentDate )}).Name } -Credential $cred  
       if( $FILEZIP ) { $log = [string] "BACKUP $server : Arquivo $ARQZIP compactado com sucesso." }
       else {           $log = [string] "BACKUP $server : Erro na compactação do arquivo $arq." }
        

### OPENSSL - CRIPTOGRAFIA ###

    $ARQZIPENC = "$ARQZIP.enc" 
    
    Invoke-Command -ComputerName $allserver -ScriptBlock { CMD.EXE /C "C:\Program Files (x86)\GnuWin32\bin\openssl.exe" enc -e -aes256 -in $BKPGDRIVE\$ARQZIP -out $BKPGDRIVE\$ARQZIPENC -pass "pass:b2h4F4d8MijKh15n1Yt158Opld1" } -Credential $cred   
    $FILEZIPENC = Invoke-Command -ComputerName $allserver -ScriptBlock { (Get-ChildItem -Filter *.enc $BKPGDRIVE|Where-Object {($_.LastWriteTime -ge $CurrentDate )}).Name } -Credential $cred 
      if($FILEZIPENC){ $log1 = [string] "BACKUP $server : Arquivo $filezipenc criptografado com sucesso."  }
      else { $log1 = [string] "BACKUP $server : Erro na criptografia do arquivo $ARQZIP."  }
      
  
### GOOGLE DRIVE ###

   $IDDIRGDRIVE = [string]'0B0IqV3rVHXVvd0toSDVqUkQ3NDA'
   Invoke-Command -ComputerName $allserver -ScriptBlock { CMD.EXE /C "D:\usr\util\gdrive.exe" upload --parent $IDDIRGDRIVE $BKPGDRIVE\$ARQZIPENC } -Credential $cred
   $FIND = Invoke-Command -ComputerName $allserver -ScriptBlock { D:\usr\util\gdrive.exe list --query "name contains '$ARQZIPENC'" } -Credential $cred
   $log2 = [string] $FIND
 
### ARQUIVOS COMPACTADOS OBSOLETOS ### 

   Invoke-Command -ComputerName $allserver -ScriptBlock {  } -Credential $cred  

   $FILEOLD = Invoke-Command -ComputerName $allserver -ScriptBlock {  Get-ChildItem $BKPGDRIVE -Recurse |Where-Object {($_.LastWriteTime -lt $DatetoDelete) -and (($_.Extension -eq ".zip") -or ($_.Extension -eq ".enc"))}|select Name,LastWriteTime |Format-Table -AutoSize |Out-String -Width 4096 } -Credential $cred  
   foreach ($FILEOLD in $FILEOLD){
        if($FILEOLD){	
                Invoke-Command -ComputerName $allserver -ScriptBlock { dir $BKPGDRIVE |Where-Object {($_.LastWriteTime -lt $DatetoDelete) -and (($_.Extension -eq ".zip") -or ($_.Extension -eq ".enc")) } | Remove-Item  } -Credential $cred  
                $log3 = [string] "BACKUP $server - Deletar Backups Compactados com mais $Daysback (dias)" }
        else {  $log3 = [string] "BACKUP $server - Não há Backups Compactados com mais $Daysback (dias)" }
	}


# ENVIO EMAIL #

  $emailTo = "marcos.cianci@s4bdigital.net"
  $subject = "Report -  Envio Backup Google Drive - $server"
  $body_rel =  "

 <HEAD>
 <STYLE type=text/css>
  BODY {text-align: left}
  p {font-size: 14px; font-family: Arial} </STYLE>
  
<p><H4> GOOGLE DRIVE - COMPACTAR E ENVIAR BACKUP SERVIDOR: $server</H4></p>
<p></p>
<p></p>
<p> Inicio: Compactar Arquivos de Backup </p>
<h4><p style=color:red > $log</p></h4>
<p> Fim: Compactar Arquivos de Backup </p>
<p></p>
<p></p>
<p> Inicio: Envio Arquivo Criptografado Google Drive </p>
<h4><p style=color:red >  $log1  </p></h4>
<p> Fim: envio Arquivos Backup Google Drive </p>
<p></p>
<p></p>
<p> Inicio: Envio Arquivos Backup Google Drive </p>
<h4><p style=color:red >  $log2  </p></h4>
<p> Fim: envio Arquivos Backup Google Drive </p>
<p></p>
<p></p>
<p> Inicio: Arquivos Obsoletos Deletados </p>
<h4><p style=color:red >  $log3  </p></h4>
<p> Fim: Arquivos Obsoletos Deletados </p>
<p></p>
<p></p>
<p>Atenciosamente,</p>
<h4><p>Equipe Service Desk - S4BDigital</p></h4>
<img src=$img alt='Image Error' vspace='10x' >

<h4><p>*** Este e-mail é gerado automaticamente - Por favor, não responda ***</p></h4>

 </BODY>
"

    EnviaEmailUser -Destino $emailTo -Mensagem $body_rel -Assunto $subject 


}

#### END ####