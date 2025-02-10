Install-Module -Name NTFSSecurity -Force
Import-Module NTFSSecurity
Import-Module ActiveDirectory
Import-Module NTFSSecurity
Clear

# Variables globales
$pathusers_protec = "c:\protec"
$pathusers_services = "c:\protec\services\"
$pathusers_employe = "c:\protec\employe\"
$DN_services = "OU=service,DC=protec-groupe,DC=com"
$Domaine = "@protec-groupe.com"
$expirationDate = Get-Date "2085-12-01"

# Création des répertoires s'ils n'existent pas
if (-not (Test-Path $pathusers_protec)) {
    New-Item -ItemType Directory -Path $pathusers_protec
}
if (-not (Test-Path $pathusers_services)) {
    New-Item -ItemType Directory -Path $pathusers_services
}
if (-not (Test-Path $pathusers_employe)) {
    New-Item -ItemType Directory -Path $pathusers_employe
}

# Lecture du fichier source
$users = Import-Csv -Path "C:\powershell\comptes-protec.csv" -Delimiter ";"
$script = $profilePath

# Liste des utilisateurs actuels
$existingUsers = Get-ADUser -Filter * -SearchBase $DN_services | Select-Object -ExpandProperty SamAccountName

# Création des utilisateurs
foreach ($user in $users) {
    $nom = $user.NOM
    $prenom = $user.PRENOM
    $id = $user.IDENTIFIANT
    $displayname = $prenom + " " + $nom
    $login = $user.IDENTIFIANT + $Domaine
    $script = $user.Script
    $no_chpassword = $false
    $groupe = $user.service
    $groupe_g = $groupe + "_g"
    $dn_classe = "OU=$groupe,OU=service,DC=protec-groupe,DC=com"
    $cn_user = "CN=" + $id + "," + $dn_classe

    # Création de l'OU service si elle n'existe pas
    if ((Get-ADOrganizationalUnit -Filter {Name -eq 'service'}) -eq $null) {
        New-ADOrganizationalUnit -Name service -Path "DC=protec-groupe,DC=com"
        Start-Sleep -Seconds 5
    }

    # Vérification si l'OU du groupe existe et la créer si nécessaire
    if ((Get-ADOrganizationalUnit -Filter {Name -eq $groupe}) -eq $null) {
        New-ADOrganizationalUnit -Name $groupe -Path "OU=service,DC=protec-groupe,DC=com"
        Start-Sleep -Seconds 5
    }

    # Vérifier si le groupe existe et le créer si nécessaire
    if ((Get-ADGroup -Filter {Name -eq $groupe_g}) -eq $null) {
        New-ADGroup -Name $groupe_g -GroupScope Global -Path "OU=service,DC=protec-groupe,DC=com"
        Start-Sleep -Seconds 5
    }

    # Gestion des partages et UO services
    Set-Location $pathusers_services
    if (-not (Test-Path $user.service)) {
        New-Item -ItemType Directory -Name $user.service
    }
    Set-Location $pathusers_employe
    if (-not (Test-Path $groupe)) {
        New-Item -ItemType Directory -Name $groupe
    }
    Set-Location "c:\protec\employe\$groupe"
    if (-not (Test-Path $id)) {
        New-Item -ItemType Directory -Name $id
    }

    # Vérification et création de l'utilisateur
    if ((Get-ADUser -Filter {SamAccountName -eq $id}) -eq $null) {
        try {
            New-ADUser -Name $displayname `
                        -GivenName $prenom `
                        -Surname $nom `
                        -SamAccountName $id `
                        -UserPrincipalName $login `
                        -Path $dn_classe `
                        -AccountPassword (ConvertTo-SecureString $user.PASSWORD -AsPlainText -Force) `
                        -DisplayName $displayname `
                        -EmailAddress $user.MESSAGERIE `
                        -Enabled $true `
                        -PasswordNeverExpires $true `
                        -ScriptPath $script `

            # Pause pour permettre la propagation de la création de l'utilisateur
            Start-Sleep -Seconds 2

            # Vérification de la création de l'utilisateur
            if (Get-ADUser -Filter {SamAccountName -eq $id}) {
                # Ajouter l'utilisateur dans le groupe service
                Add-ADGroupMember -Identity $groupe_g -Members $id
                # Définir la date d'expiration du compte
                Set-ADUser -Identity $id -AccountExpirationDate $expirationDate
                # Empêcher le changement de mot de passe
                Set-ADUser -Identity $id -CannotChangePassword $true

                # Définition des permissions NTFS pour les services
                $servicePath = Join-Path -Path $pathusers_services -ChildPath $groupe
                Add-NTFSAccess -Path $servicePath -Account $groupe_g -AccessRights 'ReadAndExecute' -AppliesTo ThisFolderSubfoldersAndFiles

                # Définition des permissions NTFS pour le dossier personnel
                $employePath = Join-Path -Path "c:\protec\employe\$groupe" -ChildPath $id
                Add-NTFSAccess -Path $employePath -Account $id -AccessRights 'Modify' -AppliesTo ThisFolderSubfoldersAndFiles

                Write-Host "L'utilisateur : $displayname est créé avec les droits NTFS appropriés !" -ForegroundColor Green
            } else {
                Write-Host "Erreur lors de la création de l'utilisateur : $displayname" -ForegroundColor Red
            }
        } catch {
            Write-Host "Erreur lors de la création de l'utilisateur : $displayname - $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "L'utilisateur : $displayname existe déjà" -ForegroundColor Magenta
    }
}

# Suppression des utilisateurs qui ne sont plus dans le fichier CSV
foreach ($existingUser in $existingUsers) {
    if (-not ($users | Where-Object { $_.IDENTIFIANT -eq $existingUser })) {
        # Suppression du compte utilisateur
        Remove-ADUser -Identity $existingUser -Confirm:$false
        # Suppression du répertoire personnel
        $personalPath = Join-Path -Path "c:\protec\employe" -ChildPath $existingUser
        if (Test-Path $personalPath) {
            Remove-Item -Path $personalPath -Recurse -Force
        }
        Write-Host "L'utilisateur : $existingUser et son répertoire personnel ont été supprimés." -ForegroundColor Cyan
    }
}