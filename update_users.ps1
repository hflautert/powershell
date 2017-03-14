# Le dados gerados em um arquivo csv e atualiza os dados no AD
# hflautert@gmail.com

. \\sabia\AREA2\DEGTI\Sinc_AD\Fontes\Logging_Functions.ps1
# http://9to5it.com/powershell-logging-function-library/

Import-Module ActiveDirectory

$data_log=Get-Date -format yyyyMMddHHmm
$pasta_de_logs="\\sabia\AREA2\DEGTI\Sinc_AD\Logs"
$pasta_de_exp="\\sabia\AREA2\DEGTI\Sinc_AD\ExpOracle"
$log="$pasta_de_logs\Andamento_script$data_log.log"
$log_movidos="$pasta_de_logs\Usuarios_movidos$data_log.log"
$log_criados="$pasta_de_logs\Usuarios_criados$data_log.log"
$log_desativados="$pasta_de_logs\Usuarios_desativados$data_log.log"
$csv_oracle="$pasta_de_exp\base_oracle_a.csv"
$csv_oracle_rd="$pasta_de_exp\base_oracle_rd.csv"
$total_processados=0
$total_atualizados=0
$total_movidos=0
$total_criados=0
$total_desativados=0

# Dados email
$mFrom="no-reply@fimfim.epagri.sc.gov.br"
#$mTo="redes@epagri.sc.gov.br"
$mTo="hflautert@gmail.com"
$mCc="henrique.lautert@ilhaservice.com.br"
$mSubject="Resultado da sincronização Oracle-AD"
$mBody=""

Log-Start -LogPath $pasta_de_logs -LogName "Andamento_script$data_log.log" -ScriptVersion "1.0"
Log-Start -LogPath $pasta_de_logs -LogName "Usuarios_criados$data_log.log" -ScriptVersion "1.0"
Log-Start -LogPath $pasta_de_logs -LogName "Usuarios_movidos$data_log.log" -ScriptVersion "1.0"
Log-Start -LogPath $pasta_de_logs -LogName "Usuarios_desativados$data_log.log" -ScriptVersion "1.0"



If (Test-Path $csv_oracle) {
    # Conversão para UTF-8
    $temp_oracle=Get-Content $csv_oracle
    $temp_oracle | Out-File -Encoding "UTF8" $csv_oracle
    $users = Import-Csv -Path $csv_oracle
    Log-Write -LogPath $log -LineValue "Arquivo: $csv_oracle carregado com sucesso."
    Log-Write -LogPath $log -LineValue "."


$users | ForEach-Object {
    #Troca ; por , para fazer consulta
    $dn_oracle=$_.DistinguishedName | %{$_ -replace ';', ','}
    $ora_ou=echo $dn_oracle | %{$_.split(",",2)[-1]}
            
    # Testa se existe usuário
    $username=$_.SamAccountName
    $givenname=$_.GivenName
    $surname=$_.Surname
    $fullname=$_.GivenName+" "+$_.Surname
    $cpf=$_.Info
    $upn=$_.SamAccountName+"@epagri.sc.gov.br"
    
    
    Log-Write -LogPath $log -LineValue "Processando usuário:  $fullname."
    Log-Write -LogPath $log -LineValue "Username: $username."

    $user=Get-ADUser $_.SamAccountName
    If ($?) {
        # Se existe confere OU, se não está na mesma, move para que veio informada pelo Oracle
        # Corta somente o endereço da OU para mover com o Move-ADObject        
        $ad_ou=($user.DistinguishedName -split ",", 2)[1]
        Log-Write -LogPath $log -LineValue "Usuario encontrado no AD, verificando posição na OU."
        If ($ad_ou -ne $ora_ou) {
            Log-Write -LogPath $log -LineValue "Usuário não está na OU enviada pelo oracle."
            Log-Write -LogPath $log_movidos -LineValue "Usuário: $fullname"
            Log-Write -LogPath $log_movidos -LineValue "Username: $username"
            Log-Write -LogPath $log_movidos -LineValue "Movendo de:$ad_ou."
            Log-Write -LogPath $log_movidos -LineValue "Para:$ora_ou."
            Log-Write -LogPath $log_movidos -LineValue "."
            Log-Write -LogPath $log -LineValue "Movendo de:$ad_ou."
            Log-Write -LogPath $log -LineValue "Para:$ora_ou."
            Get-ADUser $_.SamAccountName | Move-ADObject -TargetPath $ora_ou
            If ($?) {
                $total_movidos++
                Log-Write -LogPath $log -LineValue "Usuário movido com sucesso."
            }
            Else{
                Log-Write -LogPath $log -LineValue "Erro ao mover usuário."
            }
        }
    }
    Else {
        #Se não existe, cria com o básico, depois atualiza com campos extras.
        Log-Write -LogPath $log -LineValue "Usuário não existe no AD, criando..."
        Log-Write -LogPath $log -LineValue "Username: $username"
        Log-Write -LogPath $log -LineValue "Senha inicial: Ep$cpf"
        Log-Write -LogPath $log -LineValue "Ou: $ora_ou"
        Log-Write -LogPath $log_criados -LineValue "Usário: $fullname"
        Log-Write -LogPath $log_criados -LineValue "Username: $username"
        Log-Write -LogPath $log_criados -LineValue "Senha inicial: Ep$cpf"
        Log-Write -LogPath $log_criados -LineValue "Ou: $ora_ou"
        Log-Write -LogPath $log_criados -LineValue "."
        New-ADUser -SamAccountName $_.SamAccountName `
                   -Name $fullname `
                   -DisplayName $fullname `
                   -UserPrincipalName $upn `
                   -enable $True `
                   -accountPassword (ConvertTo-SecureString -AsPlainText "Ep$cpf" -Force) `
                   -Path $ora_ou
        If($?) {
            Log-Write -LogPath $log -LineValue "Usuário criado com sucesso."
            $total_criados++
        }
        Else {
            Log-Write -LogPath $log -LineValue "Erro ao criar usuário."
        }
    }

    Log-Write -LogPath $log -LineValue "Executando atualização de dados."
    #Alterado Departament x Description - Devido a visualização no AD e campo do OTRS
    Set-ADUSer -Identity $_.SamAccountName `
               -GivenName $_.GivenName `
               -Surname $_.Surname `
               -EmailAddress $_.EmailAddress `
               -Company $_.Company `
               -Title $_.Title `
               -Department $_.Description `
               -Description $_.Department `
               -Fax $_.Fax `
               -OfficePhone $_.OfficePhone `
               -Enabled $true `
               -Replace @{info=$_.Info; physicalDeliveryOfficeName=$_.physicalDeliveryOfficeName; pager=$_.pager}
    If ($?) {
        Log-Write -LogPath $log -LineValue "Atualização realizada com sucesso."
        $total_atualizados++
    }
    Else {
        Log-Write -LogPath $log -LineValue "Erro ao atualizar dados."
    }

Log-Write -LogPath $log -LineValue "."
$total_processados++
} # Fecha laço users
} # Fecha If carregar arquivo
Else {
    Log-Write -LogPath $log -LineValue "Erro ao carregar o arquivo, script não executado."
}

### INICIA DESATIVAÇÃO DE USUÁRIOS
Log-Write -LogPath $log -LineValue "***"
Log-Write -LogPath $log -LineValue "Inciando processo de desativação dos usuários."
Log-Write -LogPath $log -LineValue "."
If (Test-Path $csv_oracle_rd) {
    # Conversão para UTF-8
    $temp_oracle=Get-Content $csv_oracle_rd
    $temp_oracle | Out-File -Encoding "UTF8" $csv_oracle_rd
    $users = Import-Csv -Path $csv_oracle_rd
    Log-Write -LogPath $log -LineValue "Arquivo: $csv_oracle_rd carregado com sucesso."
    Log-Write -LogPath $log -LineValue "."

    $users | ForEach-Object {
    
    $username=$_.SamAccountName
    $givenname=$_.GivenName
    $surname=$_.Surname
    $fullname=$_.GivenName+" "+$_.Surname
    $data=$_.DateOff

    Log-Write -LogPath $log_desativados -LineValue "Usuário: $fullname"
    Log-Write -LogPath $log_desativados -LineValue "Username: $username"
    Log-Write -LogPath $log_desativados -LineValue "Data de desligamento: $data"
       
    Set-ADAccountExpiration $_.SamAccountName -DateTime "$data"
    Set-ADUser $_.SamAccountName -Enabled $false
    If($?){
        Log-Write -LogPath $log_desativados -LineValue "Desativado com sucesso."
        $total_desativados++
    }
    Else {
        Log-Write -LogPath $log_desativados -LineValue "Erro ao desativar usuário."
    }
    Log-Write -LogPath $log_desativados -LineValue "."
    }# Fecha laço users rd
}# Fecha laço carregar arquivo rd 
Else {
    Log-Write -LogPath $log -LineValue "Erro ao carregar o arquivo, script não executado."
}
### FIM DESATIVAÇÃO


# Compondo email
$mBody=echo "Usuários processados: $total_processados `r`n"
$mBody=$mBody+"Usuários atualizados: $total_atualizados `r`n"
$mBody=$mBody+"`n"
$mBody=$mBody+"Usuários movidos: $total_movidos `r`n"
$mBody=$mBody+"Usuários criados: $total_criados `r`n"
$mBody=$mBody+"Usuários desativados: $total_desativados `r`n"
$mBody=$mBody+"`n"
$mBody=$mBody+"Logs:`r`n"
$mBody=$mBody+"Criados:`r`n $log_criados `r`n"
$mBody=$mBody+"Movidos:`r`n $log_movidos `r`n"
$mBody=$mBody+"Desativados:`r`n $log_desativados `r`n"
$mBody=$mBody+"Andamento geral:`r`n $log `r`n"

#Send-MailMessage -SmtpServer smtp.epagri.sc.gov.br -Subject $mSubject -Body $mBody -From $mFrom -To $mTo -Cc $mCc -Encoding UTF8

Log-Finish -LogPath $log -NoExit $True
Log-Finish -LogPath $log_criados -NoExit $True
Log-Finish -LogPath $log_movidos -NoExit $True
Log-Finish -LogPath $log_desativados -NoExit $True

# Cleanning times
# Limpeza de arquivos antigos
get-childitem $pasta_de_logs | where -FilterScript {$_.LastWriteTime -le [System.DateTime]::Now.AddDays(-7)} | remove-item
get-childitem $pasta_de_exp | where -FilterScript {$_.LastWriteTime -le [System.DateTime]::Now.AddDays(-7)} | remove-item
# Renomar para manter histórico
Move-Item $csv_oracle $pasta_de_exp\base_oracle_a$data_log.csv
Move-Item $csv_oracle_rd $pasta_de_exp\base_oracle_rd$data_log.csv
