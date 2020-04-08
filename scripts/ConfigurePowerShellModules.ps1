<#
    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You
    a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of
    the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which
    the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;
    and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneysâ€™
    fees, that arise or result from the use or distribution of the Sample Code.
    Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within
    the Premier Customer Services Description.
#>

<#
    This script configures PowerShell for running the Firstline Worker bulk add tool.  It elevates the execution policy for the current
    process.
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

Function Register-Module
{
    [CmdletBinding()]
    param(
    [Parameter(Mandatory)]
    [String]
    $moduleName,

    [Parameter(Mandatory)]
    [String]
    $minVersion,

    [Parameter()]
    [String]
    $repositorySource
    )

    # Install of PowershellGet to current version
    if (-not(Get-InstalledModule -Name $moduleName -MinimumVersion $minVersion -ErrorAction SilentlyContinue))
    {
        Write-Host "Installing $moduleName Module..."
        if ($null -eq $repositorySource)
        {
            Install-Module -Name $moduleName -MinimumVersion $minVersion -Force -Scope CurrentUser -AllowClobber -AcceptLicense
        }
        else {
            Install-Module -Name $moduleName -Repository $repositorySource -MinimumVersion $minVersion -Force -Scope CurrentUser -AllowClobber -AcceptLicense
        }
    }
    else
    {
        # Check if PowershellGet module is imported in current session
        if (-not(Get-Module $moduleName -ErrorAction SilentlyContinue))
        {
            Write-Host "Importing $moduleName Module into current session..."
            Import-Module $moduleName -DisableNameChecking
        }
    }
    Write-Host "$moduleName module loaded"
}
# Set the PowerShell execution policy for the current process/session so the script is not blocked

Write-Host 'Configuring PowerShell Modules'
set-ExecutionPolicy -Scope Process -ExecutionPolicy Unrestricted -Force

# Set versions minimums for modules
$NuGetVersion = "1.3.3"
$pshGetVersion = "2.2.3"
$msalVersion = "4.10.0.1"
$azureADVersion = "2.0.2.76"
$azureADPreviewVersion = "2.0.2.89"
$cloudUtilsVersion = "1.0.0.1"
# $teamsVersion = "1.0.5"
$teamspreviewVersion = "1.0.21"
$skypeonlineVersion = "7.0.0.0"
$msonlineVersion = "1.1.183.57"

# Required to install the preview of the Microsoft Teams module
if (-not(Get-PSRepository -name PSTestGallery)) {
    Write-Host "Registering PoshTestGallery repository"
    Register-PSRepository -SourceLocation https://www.poshtestgallery.com/api/v2 -Name PSTestGallery -InstallationPolicy Trusted
}

# must install and run the SkypeOnlineConnector.exe from:
# https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.exe
Write-Host
Write-Host "****************************************************************************************************************"
Write-Host "NOTE!  The SkypeOnlineConnector must be installed locally before running this script."
Write-Host "    https://download.microsoft.com/download/2/0/5/2050B39B-4DA5-48E0-B768-583533B42C3B/SkypeOnlinePowerShell.exe"
Write-Host "****************************************************************************************************************"
Write-Host
$YesOrNo = Read-Host "DO YOU WANT TO CONTINUE - Please enter your response (y/n)"
while("y","n" -notcontains $YesOrNo )
{
    $YesOrNo = Read-Host "DO YOU WANT TO CONTINUE - Please enter your response (y/n)"
}
if (($YesOrNo).ToUpper() -eq 'n') {
    exit 0
}

Register-Module "NuGet" "$NuGetVersion" "PSGallery"
Register-Module "PowerShellGet" "$pshGetVersion" "PSGallery"
Register-Module "MSAL.PS" "$msalVersion" "PSGallery"
Register-Module "MSOnline" "$msonlineVersion" "PSGallery"
Register-Module "AzureAD" "$azureADVersion" "PSGallery"
Register-Module "AzureADPreview" "$AzureADPreviewVersion" "PSGallery"
Register-Module "MSCloudIdUtils" "$cloudUtilsVersion" "PSGallery"
# Script is using the Microsoft Teams preview module.  Commented out for clarity.
# Register-Module "MicrosoftTeams" "$teamsVersion" "PSGallery"
Register-Module "MicrosoftTeams" "$teamspreviewVersion" "PSTestGallery"

# Special case to handle SkypeOnlineConnector which was installed via the .exe from above and must be imported locally
if (-not(Get-Module -Name SkypeOnlineConnector)) {
    Import-Module "C:\\Program Files\\Common Files\\Skype for Business Online\\Modules\\SkypeOnlineConnector\\SkypeOnlineConnector.psd1" -MinimumVersion "$skypeonlineVersion"
}
Write-Host "SkypeOnlineConnector module loaded"

Write-Host "Configuring Modules and Setting Environment Completed"
Get-Module
Write-Host "Completed Configure PowerShell Modules"
