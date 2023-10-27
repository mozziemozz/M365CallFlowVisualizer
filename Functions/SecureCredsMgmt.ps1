<#

.SYNOPSIS
    This script defines functions for managing encrypted passwords and retrieving secure credentials.

    .DESCRIPTION

    Author:             Martin Heusser
    Version:            1.0.0
    Sponsor Project:    https://github.com/sponsors/mozziemozz
    Website:            https://heusser.pro

    The script contains two functions:
    - New-MZZEncryptedPassword: Prompts the user to enter a password, hashes it, and stores the encrypted password in a file.
    - Get-MZZSecureCreds: Retrieves the stored encrypted password or credentials based on the provided parameters.

    .PARAMETER FileName
    Specifies the name of the file to store the encrypted password. If not provided, the stored password will be used to create a PS credential object.

    .PARAMETER AdminUser
    Specifies the username for which the credentials are stored. If not provided, $env:USERNAME will be used.

    .PARAMETER checkPassword
    Indicates whether to display the decrypted password after retrieving the credentials.

    .PARAMETER updatePassword
    Specifies whether to update the stored password or credentials.

    .NOTES
    - This script requires Git to be installed and accessible in the environment for retrieving the repository path.

    .EXAMPLE
    New-MZZEncryptedPassword -FileName "MyPassword"

    This example prompts the user to enter a password and stores the encrypted password in the "MyPassword.txt" file.

    .EXAMPLE
    Get-MZZSecureCreds -FileName "MyPassword" -CheckPassword

    This example retrieves the encrypted password from the "MyPassword.txt" file, decrypts it, and displays the decrypted password.

    .EXAMPLE
    New-MZZEncryptedPassword -AdminUser "admin@domain.com"

    This example prompts the user to enter a password and stores the encrypted password in the "admin@domain.com" file.

    .EXAMPLE
    Get-MZZSecureCreds -AdminUser "admin@domain.com"

    This example retrieves the encrypted password from the "admin@domain.com" file, as part of a PS credential object.


#>

function New-MZZEncryptedPassword {
    param (
        [Parameter(Mandatory = $false)][string]$FileName,
        [Parameter(Mandatory = $false)][string]$AdminUser = $env:USERNAME
    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }
    
    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    $SecureStringPassword = Read-Host "Please enter the password you would like to hash" -AsSecureString
    
    $PasswordHash = $SecureStringPassword | ConvertFrom-SecureString

    if ($FileName) {

        Set-Content -Path "$secureCredsFolder\$($FileName).txt" -Value $PasswordHash -Force

        Get-MZZSecureCreds -fileName $FileName

    }

    else {

        Set-Content -Path "$secureCredsFolder\$($adminUser).txt" -Value $PasswordHash -Force

        Get-MZZSecureCreds -AdminUser $adminUser

    }

}

function Get-MZZSecureCreds {
    param (
        [Parameter(Mandatory = $false)][switch]$CheckPassword,
        [Parameter(Mandatory = $false)][switch]$UpdatePassword,
        [Parameter(Mandatory = $false)][string]$FileName,
        [Parameter(Mandatory = $false)][string]$AdminUser = $env:USERNAME,
        [Parameter(Mandatory = $false)][switch]$NoClipboard

    )

    if (!$localRepoPath) {

        $localRepoPath = git rev-parse --show-toplevel

    }

    $secureCredsFolder = "$localRepoPath\.local\SecureCreds"

    if (!(Test-Path -Path $secureCredsFolder)) {

        New-Item -Path $secureCredsFolder -ItemType Directory

    }

    if ($FileName) {

        if ($UpdatePassword) {

            New-MZZEncryptedPassword -fileName $FileName

        }

        if (!(Test-Path -Path "$localRepoPath\.local\SecureCreds\$FileName.txt")) {

            Write-Host "No password found for filename: $FileName..." -ForegroundColor Yellow

            New-MZZEncryptedPassword -fileName $FileName

        }

        else {

            $passwordEncrypted = Get-Content -Path "$localRepoPath\.local\SecureCreds\$($FileName).txt" | ConvertTo-SecureString

            if (!$passwordEncrypted) {

                New-MZZEncryptedPassword -fileName $FileName

            }

            $global:passwordDecrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($passwordEncrypted))
            
            if ($CheckPassword) {

                Write-Host "Decrypted password: $passwordDecrypted" -ForegroundColor Cyan

            }

        }

        if (!$NoClipboard) {

            Write-Host "Password is stored in `$passwordDecrypted variable and in clipboard!" -ForegroundColor Yellow

            $passwordDecrypted | Set-Clipboard

        }

        return $passwordDecrypted

    }

    else {

        if ($UpdatePassword) {

            New-MZZEncryptedPassword -AdminUser $adminUser

        }


        if (!(Test-Path -Path "$LocalRepoPath\.local\SecureCreds\$adminUser.txt")) {

            Write-Host "No credentials found for user: $adminUser..." -ForegroundColor Yellow

            New-MZZEncryptedPassword -AdminUser $adminUser

        }

        else {

            $adminPasswordEncrypted = Get-Content -Path "$localRepoPath\.local\SecureCreds\$($adminUser).txt" | ConvertTo-SecureString

            if (!$adminPasswordEncrypted) {

                New-MZZEncryptedPassword

            }

            Write-Host $adminUser -ForegroundColor Green

            $global:secureCreds = New-Object System.Management.Automation.PSCredential -ArgumentList $adminUser,$adminPasswordEncrypted

            if ($CheckPassword) {

                $adminPasswordDecrypted = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureCreds.Password))
                Write-Host "Decrypted password: $adminPasswordDecrypted" -ForegroundColor Cyan

            }

        }

        Write-Host "Credentials are stored in `$secureCreds variable!" -ForegroundColor Cyan

        return $secureCreds > $null

    }

}

function Get-MZZTenantIdTxt {
    param (
    )

    if (!(Test-Path -Path .\.local\SecureCreds\TenantId.txt)) {

        $TenantId = Read-Host "Enter your Tenant Id"

        Set-Content -Path .\.local\SecureCreds\TenantId.txt -Value $TenantId

    }

    else {

        $TenantId = (Get-Content -Path .\.local\SecureCreds\TenantId.txt).Trim()

    }
    
}

function Get-MZZAppIdTxt {
    param (
    )

    if (!(Test-Path -Path .\.local\SecureCreds\AppId.txt)) {

        $AppId = Read-Host "Enter your App Id"

        Set-Content -Path .\.local\SecureCreds\AppId.txt -Value $AppId

    }

    else {

        $AppId = (Get-Content -Path .\.local\SecureCreds\AppId.txt).Trim()

    }
    
}