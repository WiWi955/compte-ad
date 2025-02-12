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
# SIG # Begin signature block
# MIIcwAYJKoZIhvcNAQcCoIIcsTCCHK0CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBZBECGi+rrm4Gr
# wGZupIzqSlEXQN5wsRGwBHDCPuz5SqCCAwQwggMAMIIB6KADAgECAhBqfG1qTM7n
# mkUzNfb85kGVMA0GCSqGSIb3DQEBCwUAMBgxFjAUBgNVBAMMDU5ld0NlcnRpZmlj
# YXQwHhcNMjUwMjExMTIwNzI1WhcNMjYwMjExMTIyNzI1WjAYMRYwFAYDVQQDDA1O
# ZXdDZXJ0aWZpY2F0MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAk2E1
# VL5k/x/TqoEeYMUi0O1kFrfx+rPNbck43W3YfhPS4TBKlKko5TvgxEKbshptuEXH
# vvcQApFKpa9ouG1PEbVDbA2+xuy1navWK1NFlj5mzqLIwr7NdMLUnUu0fOaQ8xJo
# qWKTdUl75paIPeurojCTQU687tIivZ4wf5N2lw0CySSV+hMdSXmOgWFed2UiZ/n8
# fyu1uf1Z82j1Z4EyijylB+5x9nryzf+IrXC6be+ZH6xpAxlz7TQcChaxpuEiLqZ+
# IhyuM9xccaKGNl6RhF4BO9DmtNmumPZoieHKYAvGto0oFIUH/L7/shDm1gtw7Zp+
# deXdXYgObOi7M2y+cQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFO0maH3i9vkr7e2HLmEBfkeCwJJkMA0GCSqG
# SIb3DQEBCwUAA4IBAQAeBtT2uIF5MSsh/SzIH2wJk/E/ZhjIUURumZ5VsOPi1phA
# y3X2lQG/SgKAJtWcltio3G9zs459RO+1zGtaCNr4qOjS2bJymw87KUeofEViMv6u
# qZHcMn0IBUnEz2sgmv8IMZcJHqxGoyYZbgIHTAsX9l7uLJ4qt01AHD4HdGofSQGG
# d2/TVQ8exgXHl4GZ68ISt10NEWMTKoIHga77ApDO3yol3+wX/8DK1mH72DpK7cC0
# 0c5Cl8KAI2He2bnTVUYdjh4wAmf9NTUAvBkSMDCPJTj/43pdMdLOYp3f9LLOg1VN
# kD6bISdEdZFmHSmLmQWpi6duLV9jXSyqv0EqeVBtMYIZEjCCGQ4CAQEwLDAYMRYw
# FAYDVQQDDA1OZXdDZXJ0aWZpY2F0AhBqfG1qTM7nmkUzNfb85kGVMA0GCWCGSAFl
# AwQCAQUAoHwwEAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQB
# gjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkE
# MSIEINdq5Av3nX8+5+mPXHMCyQ+zfmhkWcrhWmZg3g5S7uQZMA0GCSqGSIb3DQEB
# AQUABIIBAGojV9fHK13JtgNa1F4HR09Mr0QGlY6prchc0yDL2CZPfYHoJxP4qr17
# PrVkdN8OobcEJHub87MWTtFkfykRFwMVNFBimL0MA+/RegznBuLDQmFV02CVOUaG
# vDGRUrd/y1koJp/k9031/B2w/XBwc50XmZMFNetujdRfRo4wZ05+rmbgeA3KuvCA
# 9vNckWH2hTsCQlO8Z0OiLlx8SyftMlJsuSobF5dFSQPlEoC4DKdPlcsghl63gdiE
# Zocx1dV4HYDoOms50eUSraRfhIPj/r3FfzdoPJuXLkaXeQlYJyVnH+KkPFYrGTT8
# KbwZMJJlSwjfODlsbU/Vyx+Qq00lf2Ghghc5MIIXNQYKKwYBBAGCNwMDATGCFyUw
# ghchBgkqhkiG9w0BBwKgghcSMIIXDgIBAzEPMA0GCWCGSAFlAwQCAQUAMHcGCyqG
# SIb3DQEJEAEEoGgEZjBkAgEBBglghkgBhv1sBwEwMTANBglghkgBZQMEAgEFAAQg
# FOG8l6atgbRM0MklYURP9wdwhStALFLEhvn6Prn9wYACECCFSAmxnhvZwdAOyfOc
# KNYYDzIwMjUwMjExMTUwNTE5WqCCEwMwgga8MIIEpKADAgECAhALrma8Wrp/lYfG
# +ekE4zMEMA0GCSqGSIb3DQEBCwUAMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5E
# aWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0
# MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcgQ0EwHhcNMjQwOTI2MDAwMDAwWhcNMzUx
# MTI1MjM1OTU5WjBCMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNlcnQxIDAe
# BgNVBAMTF0RpZ2lDZXJ0IFRpbWVzdGFtcCAyMDI0MIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEAvmpzn/aVIauWMLpbbeZZo7Xo/ZEfGMSIO2qZ46XB/Qow
# IEMSvgjEdEZ3v4vrrTHleW1JWGErrjOL0J4L0HqVR1czSzvUQ5xF7z4IQmn7dHY7
# yijvoQ7ujm0u6yXF2v1CrzZopykD07/9fpAT4BxpT9vJoJqAsP8YuhRvflJ9YeHj
# es4fduksTHulntq9WelRWY++TFPxzZrbILRYynyEy7rS1lHQKFpXvo2GePfsMRhN
# f1F41nyEg5h7iOXv+vjX0K8RhUisfqw3TTLHj1uhS66YX2LZPxS4oaf33rp9Hlfq
# SBePejlYeEdU740GKQM7SaVSH3TbBL8R6HwX9QVpGnXPlKdE4fBIn5BBFnV+KwPx
# RNUNK6lYk2y1WSKour4hJN0SMkoaNV8hyyADiX1xuTxKaXN12HgR+8WulU2d6zhz
# XomJ2PleI9V2yfmfXSPGYanGgxzqI+ShoOGLomMd3mJt92nm7Mheng/TBeSA2z4I
# 78JpwGpTRHiT7yHqBiV2ngUIyCtd0pZ8zg3S7bk4QC4RrcnKJ3FbjyPAGogmoiZ3
# 3c1HG93Vp6lJ415ERcC7bFQMRbxqrMVANiav1k425zYyFMyLNyE1QulQSgDpW9rt
# vVcIH7WvG9sqYup9j8z9J1XqbBZPJ5XLln8mS8wWmdDLnBHXgYly/p1DhoQo5fkC
# AwEAAaOCAYswggGHMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBYGA1Ud
# JQEB/wQMMAoGCCsGAQUFBwMIMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG
# /WwHATAfBgNVHSMEGDAWgBS6FtltTYUvcyl2mi91jGogj57IbzAdBgNVHQ4EFgQU
# n1csA3cOKBWQZqVjXu5Pkh92oFswWgYDVR0fBFMwUTBPoE2gS4ZJaHR0cDovL2Ny
# bDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1NlRp
# bWVTdGFtcGluZ0NBLmNybDCBkAYIKwYBBQUHAQEEgYMwgYAwJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBYBggrBgEFBQcwAoZMaHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0UlNBNDA5NlNIQTI1
# NlRpbWVTdGFtcGluZ0NBLmNydDANBgkqhkiG9w0BAQsFAAOCAgEAPa0eH3aZW+M4
# hBJH2UOR9hHbm04IHdEoT8/T3HuBSyZeq3jSi5GXeWP7xCKhVireKCnCs+8GZl2u
# VYFvQe+pPTScVJeCZSsMo1JCoZN2mMew/L4tpqVNbSpWO9QGFwfMEy60HofN6V51
# sMLMXNTLfhVqs+e8haupWiArSozyAmGH/6oMQAh078qRh6wvJNU6gnh5OruCP1QU
# AvVSu4kqVOcJVozZR5RRb/zPd++PGE3qF1P3xWvYViUJLsxtvge/mzA75oBfFZSb
# dakHJe2BVDGIGVNVjOp8sNt70+kEoMF+T6tptMUNlehSR7vM+C13v9+9ZOUKzfRU
# AYSyyEmYtsnpltD/GWX8eM70ls1V6QG/ZOB6b6Yum1HvIiulqJ1Elesj5TMHq8CW
# T/xrW7twipXTJ5/i5pkU5E16RSBAdOp12aw8IQhhA/vEbFkEiF2abhuFixUDobZa
# A0VhqAsMHOmaT3XThZDNi5U2zHKhUs5uHHdG6BoQau75KiNbh0c+hatSF+02kULk
# ftARjsyEpHKsF7u5zKRbt5oK5YGwFvgc4pEVUNytmB3BpIiowOIIuDgP5M9WArHY
# SAR16gc0dP2XdkMEP5eBsX7bf/MGN4K3HP50v/01ZHo/Z5lGLvNwQ7XHBx1yomzL
# P8lx4Q1zZKDyHcp4VQJLu2kWTsKsOqQwggauMIIElqADAgECAhAHNje3JFR82Ees
# /ShmKl5bMA0GCSqGSIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxE
# aWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMT
# GERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAz
# MjIyMzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5j
# LjE7MDkGA1UEAxMyRGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBU
# aW1lU3RhbXBpbmcgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDG
# hjUGSbPBPXJJUVXHJQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6
# ffOciQt/nR+eDzMfUBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/
# qxkrPkLcZ47qUT3w1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3Hxq
# V3rwN3mfXazL6IRktFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVj
# bOSmxR3NNg1c1eYbqMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcp
# licu9Yemj052FVUmcJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZ
# girHkr+g3uM+onP65x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZG
# s506o9UD4L/wojzKQtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHz
# NklNiyDSLFc1eSuo80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2
# ElGTyYwMO1uKIqjBJgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJ
# ASgADoRU7s7pXcheMBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYD
# VR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8w
# HwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGG
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcw
# AYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8v
# Y2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBD
# BgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRUcnVzdGVkUm9vdEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgB
# hv1sBwEwDQYJKoZIhvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4Q
# TRPPMFPOvxj7x1Bd4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfN
# thKWb8RQTGIdDAiCqBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1g
# tqpPkWaeLJ7giqzl/Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1Ypx
# dmXazPByoyP6wCeCRK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/um
# nXKvxMfBwWpx2cYTgAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+U
# zTl63f8lY5knLD0/a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhz
# q6YBT70/O3itTK37xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11
# LB4nLCbbbxV7HhmLNriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCY
# oCvtlUG3OtUVmDG0YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvk
# dgIm2fBldkKmKYcJRyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3
# OBqhK/bt1nz8MIIFjTCCBHWgAwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG
# 9w0BAQwFADBlMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkw
# FwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1
# cmVkIElEIFJvb3QgQ0EwHhcNMjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3Qg
# RzQwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAi
# MGkz7MKnJS7JIT3yithZwuEppz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnny
# yhHS5F/WBTxSD1Ifxp4VpX6+n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE
# 5nQ7bXHiLQwb7iDVySAdYyktzuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm
# 7nfISKhmV1efVFiODCu3T6cw2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5
# w3jHtrHEtWoYOAMQjdjUN6QuBX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsD
# dV14Ztk6MUSaM0C/CNdaSaTC5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1Z
# XUJ2h4mXaXpI8OCiEhtmmnTK3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS0
# 0mFt6zPZxd9LBADMfRyVw4/3IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hk
# pjPRiQfhvbfmQ6QYuKZ3AeEPlAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m8
# 00ERElvlEFDrMcXKchYiCd98THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+i
# sX4KJpn15GkvmB0t9dmpsh3lGwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB
# /zAdBgNVHQ4EFgQU7NfjgtJxXWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReui
# r/SSy4IxLVGLp6chnfNtyA8wDgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0w
# azAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUF
# BzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVk
# SURSb290Q0EuY3J0MEUGA1UdHwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2lj
# ZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAG
# BgRVHSAAMA0GCSqGSIb3DQEBDAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9
# mqyhhyzshV6pGrsi+IcaaVQi7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxS
# A8hO0Cre+i1Wz/n096wwepqLsl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/
# 6Fmo8L8vC6bp8jQ87PcDx4eo0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSM
# b++hUD38dglohJ9vytsgjTVgHAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt
# 9H5xaiNrIv8SuFQtJ37YOtnwtoeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMYIDdjCC
# A3ICAQEwdzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4x
# OzA5BgNVBAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGlt
# ZVN0YW1waW5nIENBAhALrma8Wrp/lYfG+ekE4zMEMA0GCWCGSAFlAwQCAQUAoIHR
# MBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUw
# MjExMTUwNTE5WjArBgsqhkiG9w0BCRACDDEcMBowGDAWBBTb04XuYtvSPnvk9nFI
# UIck1YZbRTAvBgkqhkiG9w0BCQQxIgQgwzbjpo15iYM/oXsLIpYz8n5AnVIJO6BL
# Jbc57wMS89QwNwYLKoZIhvcNAQkQAi8xKDAmMCQwIgQgdnafqPJjLx9DCzojMK7W
# VnX+13PbBdZluQWTmEOPmtswDQYJKoZIhvcNAQEBBQAEggIAa3lnWo3UCeKq1FkH
# 2DDRKo6Tcxon8MRT4Lj0ABmkM/fYpMBDLgNQTAjz8MMllVnbirNHPK9qkRuzPT3+
# ZHczN+mqwtEWhNU6uNfy9C7PMIEaYDKRfgRg/8Fpv1Ym9exG9y2sNXe88q89AKpY
# slRAVV4eKe0PBN2SMQj+c23eAPhuOEDinpeg5/vJFm4eL498fSjWUEq2fyiN9wlw
# sNuqIo5LCBqHTg4GAm8aPbXPMNUTcxE53ruP7doiMVAX6bvPO24O/c5GO2FSq0nz
# vO3eRyUl6gGjbwuHh68yUZYeFYGufqFuVxOFI2SYEV2HhsPcIE7UACIuNjL1M+h6
# MkFTtwDhXoVrqbuZYl1BwdsFXvIo4hDahLABiaEwFpOFFaLPBvYpz957f6M24K8q
# AIlyl9y8PTsD9mh3svrTdTx1R41wrJqmVRns/E+/RmGpmQIrRsGfnNcdzOAjCesC
# kIC2CmylYqPfhlpLeZe+ypbS50/7sunEpu79TcLzYc68jipzFC5LfRnwgxyGVT1e
# y0MFTdMwtMJVc1oeosQSxUajEugTTrmVyNzISR9XzfdKDC4hGOx2ctETXfK8utlp
# JWBIR0+mQOpKnoQ23N/3MR01gRbFz7kb6DOOgefZPN3uhnphVeFYOPXCyT4kNm2I
# Eyx3sbqkNACVIPjzqJP51jD3+IA=
# SIG # End signature block
