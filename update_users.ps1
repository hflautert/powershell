# Le dados gerados em um arquivo csv e atualiza os dados no AD
# hflautert@gmail.com

Import-Module ActiveDirectory

$users = Import-Csv -Path \\sabia\AREA2\DEGTI\Sinc_AD\teste.csv

$users | ForEach-Object {
    #Troca ; por , para fazer consulta
    $dn_oracle=$_.DistinguishedName | %{$_ -replace ';', ','}
    echo "`nAnalisando usuário:"$dn_oracle
    
    # Testa se existe usuário
    $existe=Get-ADUser $_.SamAccountName
    If ($existe -eq $Null) {
        "`nUsuario nao existe no AD, criar"
    }
    # Caso existe, confere OU, se não está na mesma, move para que veio informada pelo Oracle
    Else {
        "`nUsuario encontrado no AD, verificando posição na OU"
        $esta_nesta_ou=Get-ADUser -LDAPFilter "(distinguishedName=$dn_oracle)"
        If ($esta_nesta_ou -eq $Null) {
            "`nUsuário não está na OU enviada pelo oracle, movendo para OU:"        
            $ou=echo $dn_oracle | %{$_.split(',',2)[-1]}
            echo $ou
            Get-ADUser $_.SamAccountName | Move-ADObject -TargetPath $ou
            
        }
        
        "`nExecutando atualização de dados`n"
        Set-ADUSer -Identity $_.SamAccountName `
               -GivenName $_.GivenName `
               -EmailAddress $_.EmailAddress `
               -Company $_.Company `
               -Title $_.Title `
               -Department $_.Department `
               -Description $_.Description `
               -Fax $_.Fax `
               -OfficePhone $_.OfficePhone `
               -Replace @{info=$_.Info; physicalDeliveryOfficeName=$_.physicalDeliveryOfficeName; pager=$_.pager}
        
    }
    
}
