# Installation et importation des modules
Install-Module -Name NTFSSecurity -Force
Import-Module NTFSSecurity
Import-Module ActiveDirectory
Import-Module NTFSSecurity
Clear

# Variables globales
$pathusers_protec = "c:\protec"
$pathusers_services = "c:\protec\services"
$pathusers_employe = "c:\protec\employe"
$DN_services = "OU=service,DC=protec-groupe,DC=com"
$Domaine = "@protec-groupe.com"
$expirationDate = Get-Date "2085-12-01"
$csvPath = "C:\powershell\comptes-protec.csv"

# Création des répertoires s'ils n'existent pas
foreach ($path in @($pathusers_protec, $pathusers_services, $pathusers_employe)) {
    if (-not (Test-Path $path)) {
        New-Item -ItemType Directory -Path $path | Out-Null
    }
}

# Vérification de l'existence du fichier CSV
if (-not (Test-Path $csvPath)) {
    Write-Host "Le fichier CSV $csvPath est introuvable !" -ForegroundColor Red
    Exit
}

# Importation des utilisateurs depuis le fichier CSV
$users = Import-Csv -Path $csvPath -Delimiter ";"

# Récupération des utilisateurs existants
$existingUsers = Get-ADUser -Filter * -SearchBase $DN_services | Select-Object -ExpandProperty SamAccountName

# Création des utilisateurs
foreach ($user in $users) {
    $nom = $user.NOM
    $prenom = $user.PRENOM
    $id = $user.IDENTIFIANT
    $displayname = "$prenom $nom"
    $login = "$id$Domaine"
    $script = $user.Script
    $groupes = $user.GROUPES_AD -split ","
    $dossier = $user.DOSSIER_NTFS
    $droits = $user.DROITS_NTFS
    $dn_classe = "OU=$($user.SERVICE),OU=service,DC=protec-groupe,DC=com"

    # Vérification et création de l'OU service si nécessaire
    if (-not (Get-ADOrganizationalUnit -Filter {Name -eq $user.SERVICE})) {
        New-ADOrganizationalUnit -Name $user.SERVICE -Path "OU=service,DC=protec-groupe,DC=com"
    }

    # Vérification et création du groupe global de service
    $groupe_g = "$($user.SERVICE)_g"
    if (-not (Get-ADGroup -Filter {Name -eq $groupe_g})) {
        New-ADGroup -Name $groupe_g -GroupScope Global -Path "OU=service,DC=protec-groupe,DC=com"
    }

    # Création de l'utilisateur si inexistant
    if (-not (Get-ADUser -Filter {SamAccountName -eq $id})) {
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
                        -ScriptPath $script
            
            Start-Sleep -Seconds 2

            # Ajout au groupe de service
            Add-ADGroupMember -Identity $groupe_g -Members $id

            # Définition de l'expiration et du mot de passe
            Set-ADUser -Identity $id -AccountExpirationDate $expirationDate
            Set-ADUser -Identity $id -CannotChangePassword $true

            # Création des dossiers personnels
            $userPath = Join-Path -Path "$pathusers_employe\$($user.SERVICE)" -ChildPath $id
            if (-not (Test-Path $userPath)) { New-Item -ItemType Directory -Path $userPath }

            # Application des permissions NTFS
            if ($dossier -and (Test-Path $dossier)) {
                $regle = New-Object System.Security.AccessControl.FileSystemAccessRule($id, $droits, "ContainerInherit,ObjectInherit", "None", "Allow")
                $acl = Get-Acl -Path $dossier
                $acl.SetAccessRule($regle)
                Set-Acl -Path $dossier -AclObject $acl
                Write-Host "Permissions $droits appliquées sur $dossier pour $id" -ForegroundColor Green
            }

            Write-Host "Utilisateur $displayname créé avec succès !" -ForegroundColor Green
        } catch {
            Write-Host "Erreur création de $displayname : $_" -ForegroundColor Red
        }
    } else {
        Write-Host "L'utilisateur $displayname existe déjà." -ForegroundColor Yellow
    }
}

# Suppression des utilisateurs obsolètes
foreach ($existingUser in $existingUsers) {
    if (-not ($users | Where-Object { $_.IDENTIFIANT -eq $existingUser })) {
        Remove-ADUser -Identity $existingUser -Confirm:$false
        $personalPath = Join-Path -Path "$pathusers_employe" -ChildPath $existingUser
        if (Test-Path $personalPath) {
            Remove-Item -Path $personalPath -Recurse -Force
        }
        Write-Host "Utilisateur supprimé : $existingUser" -ForegroundColor Cyan
    }
}

Write-Host "Synchronisation terminée." -ForegroundColor Magenta
