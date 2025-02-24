# Installation et importation des modules
Install-Module -Name NTFSSecurity -Force
Import-Module NTFSSecurity
Import-Module ActiveDirectory
Import-Module NTFSSecurity
Clear

# Chemins et variables globales
$pathusers_protec    = "C:\protec"
$pathusers_services  = "C:\protec\services"
$pathusers_employe   = "C:\protec\employe"
$DN_services         = "OU=service,DC=protec-groupe,DC=com"
$Domaine             = "@protec-groupe.com"
$expirationDate      = Get-Date "2085-12-01"
$csvPath             = "C:\powershell\comptes-protec.csv"

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

# Récupération des utilisateurs existants (optionnel, ici uniquement pour vérification)
$existingUsers = Get-ADUser -Filter * -SearchBase $DN_services | Select-Object -ExpandProperty SamAccountName

# Boucle de création des utilisateurs
foreach ($user in $users) {

    # Vérification et assignation de valeurs par défaut pour chaque champ
    $nom       = if ([string]::IsNullOrEmpty($user.NOM))         { "NomNonDefini" }       else { $user.NOM }
    $prenom    = if ([string]::IsNullOrEmpty($user.PRENOM))      { "PrenomNonDefini" }    else { $user.PRENOM }
    $id        = if ([string]::IsNullOrEmpty($user.IDENTIFIANT)) { "user_" + (Get-Random) } else { $user.IDENTIFIANT }
    $password  = if ([string]::IsNullOrEmpty($user.PASSWORD))    { "MotDePasseParDefaut123!" } else { $user.PASSWORD }
    $service   = if ([string]::IsNullOrEmpty($user.Service))       { "ServiceCommun" }      else { $user.Service }
    $messagerie= $user.MESSAGERIE  # Peut rester vide si non indispensable
    $script    = if ([string]::IsNullOrEmpty($user.Script))      { $null }                else { $user.Script }
    
    $displayname = "$prenom $nom"
    $login       = "$id$Domaine"
    $dn_classe   = "OU=$service,OU=service,DC=protec-groupe,DC=com"

    # Vérification et création de l'OU correspondant au service si nécessaire
    if (-not (Get-ADOrganizationalUnit -Filter {Name -eq $service} -ErrorAction SilentlyContinue)) {
        try {
            New-ADOrganizationalUnit -Name $service -Path "OU=service,DC=protec-groupe,DC=com" -ErrorAction Stop
            Write-Host "OU '$service' créée." -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de la création de l'OU '$service' : $_" -ForegroundColor Red
        }
    }

    # Vérification et création du groupe global de service
    $groupe_g = "$service" + "_g"
    if (-not (Get-ADGroup -Filter {Name -eq $groupe_g} -ErrorAction SilentlyContinue)) {
        try {
            New-ADGroup -Name $groupe_g -GroupScope Global -Path "OU=service,DC=protec-groupe,DC=com" -ErrorAction Stop
            Write-Host "Groupe '$groupe_g' créé." -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de la création du groupe '$groupe_g' : $_" -ForegroundColor Red
        }
    }

    # Création de l'utilisateur s'il n'existe pas déjà
    if (-not (Get-ADUser -Filter {SamAccountName -eq $id} -ErrorAction SilentlyContinue)) {
        try {
            New-ADUser -Name $displayname `
                       -GivenName $prenom `
                       -Surname $nom `
                       -SamAccountName $id `
                       -UserPrincipalName $login `
                       -Path $dn_classe `
                       -AccountPassword (ConvertTo-SecureString $password -AsPlainText -Force) `
                       -DisplayName $displayname `
                       -EmailAddress $messagerie `
                       -Enabled $true `
                       -PasswordNeverExpires $true `
                       -ScriptPath $script `
                       -ErrorAction Stop
            Start-Sleep -Seconds 2

            # Ajout de l'utilisateur au groupe global de service
            try {
                Add-ADGroupMember -Identity $groupe_g -Members $id -ErrorAction Stop
                Write-Host "Utilisateur ajouté au groupe '$groupe_g'." -ForegroundColor Green
            } catch {
                Write-Host "Erreur lors de l'ajout de $id au groupe '$groupe_g' : $_" -ForegroundColor Yellow
            }

            # Définition de la date d'expiration et configuration du mot de passe
            try {
                Set-ADUser -Identity $id -AccountExpirationDate $expirationDate -ErrorAction Stop
                Set-ADUser -Identity $id -CannotChangePassword $true -ErrorAction Stop
            } catch {
                Write-Host "Erreur de configuration pour $id : $_" -ForegroundColor Yellow
            }

            # Création du dossier personnel
            $userPath = Join-Path -Path "$pathusers_employe\$service" -ChildPath $id
            if (-not (Test-Path $userPath)) {
                New-Item -ItemType Directory -Path $userPath | Out-Null
                Write-Host "Dossier personnel créé pour $id." -ForegroundColor Green
            }

            Write-Host "Utilisateur '$displayname' créé avec succès !" -ForegroundColor Green
        } catch {
            Write-Host "Erreur lors de la création de l'utilisateur '$displayname' : $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "L'utilisateur '$displayname' existe déjà." -ForegroundColor Yellow
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
