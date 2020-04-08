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
    This script creates the AAD AU, and adds members
#>

# Create Azure AD Administrative Units
Function Import-BulkAzureADAUData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $teamsCSVPath
    )
    if ($null -eq ($teamsCSV = Read-CSV $teamsCSVPath)) {
        Write-Host "ERROR - Cannot Process Filename $usersCSVPath"
        exit
    }
    else {
        [System.Array]$AzureADRoles = Get-AzureADDirectoryRole
        if ("fe930be7-5e62-47db-91af-98c3a49a38b1" -notin $AzureADRoles.RoleTemplateId) {
            Enable-AzureADDirectoryRole -RoleTemplateId "fe930be7-5e62-47db-91af-98c3a49a38b1"
        }
        $UserAccountAdministratorRole = $($AzureADRoles | Where-Object DisplayName -eq "User Account Administrator")
        [System.Array]$AzureADAUList = Get-AzureADAdministrativeUnit
        ForEach ($team in $teamsCSV) {
            try {
                if ($team.LocationAU -notin $AzureADAUList.DisplayName) {
                    New-AzureADAdministrativeUnit -Description $team.TeamDescription -DisplayName $team.LocationAU
                }
            }
            catch {
                Write-Host "ERROR creating Azure AD Administrative Unit"
                $Error[0] | Format-List * -Force
            }


            # $AzureADAU = $null; $AzureADAU = Get-AzureADAdministrativeUnit | Where-Object DisplayName -eq $team.LocationAU
            # [System.Array]$LocationAdmins = $team.LocationAdmins.Split(";").Trim()
            # ForEach ($admin in $LocationAdmins) {
            #     $AdminUser = $null; $AdminUser = Get-AzureADUser | Where-Object UserPrincipalName -eq $admin
            #     if ($AdminUser.Count -eq 1) {
            #         $RoleMemberInfo = New-Object -TypeName Microsoft.Open.AzureAD.Model.RoleMemberInfo -Property @{ObjectId = $AdminUser.ObjectId }
            #         Add-AzureADScopedRoleMembership -RoleObjectId $UserAccountAdministratorRole.ObjectId -ObjectId $AzureADAU.ObjectId -RoleMemberInfo $RoleMemberInfo
            #     }
            #     else {
            #         $obj = New-Object PSObject
            #         $obj | Add-Member Noteproperty -Name Exception -value "Messaging Policy Already Exists (skipping) - $policyName"
            #         $obj | Add-Member Noteproperty -Name Message -value "Account Admin delegates $admin was found $($AdminUser.Count) times"
            #         $obj | Add-Member Noteproperty -Name Cmdlet -value "ConfigureAdministrativeUnits.ps1 Import-BulkAzureADAUData"
            #         LogException "$rootpath\logs\CreateTeamsMessagingPolicyExceptions.csv" $obj
            #         throw "Account Admin delegates $admin was found $($AdminUser.Count) times"
            #     }
            # }
        }
    }
}

# Add Azure AD Administrative Unit members
function Add-BulkAzureADAUMemberData {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]
        $usersCSVPath
    )
    if ($null -eq ($usersCSV = Read-CSV $usersCSVPath)) {
        Write-Host "ERROR - Cannot Process Filename $usersCSVPath"
        return
    }
    else {
        foreach ($user in $usersCSV) {
            try {
                $teamNameInfo = $user.TeamName.Split();
                $userStoreName = "Store"+$teamNameInfo[-1]+"AU"
                $userEmail = $user.Name+"@"+$tenantName
                $storeAU = Get-AzureADAdministrativeUnit | Where-Object DisplayName -eq $userStoreName
                $flwUser = Get-AzureADUser | Where-Object UserPrincipalName -eq $userEmail
                if (Get-AzureADAdministrativeUnitMember -ObjectId $storeAU.ObjectId | Where-Object ObjectId -Contains $flwUser.ObjectId) {
                    Write-Host "$($flwUser.UserPrincipalName) is already a member of the AU $($user.StoreAU)"
                    Set-Exception -FileName "$rootPath\logs\ConfigureAdministrativeUnits.csv" -Module "ConfigureAdministrativeUnits" -Message "$($flwUser.UserPrincipalName) is already a member of the AU $($user.StoreAU)"
                }
                else {
                    Add-AzureADAdministrativeUnitMember -ObjectId $storeAU.ObjectId -RefObjectId $flwUser.ObjectId
                }
            }
            catch {
                Write-Host "ERROR adding users to Azure AD Administrative Unit"
                $Error[0] | Format-List * -Force
            }
        }
    }
}

# Check Azure AD Administrative Unit membership
function Get-BulkAzureADAUMembership {
    [CmdletBinding()]
    param ()
    Get-AzureADAdministrativeUnit | % {
        $au = $_.DisplayName
        Get-AzureADAdministrativeUnitMember -ObjectId $_.ObjectId | % {
            $user = Get-AzureADUser -ObjectId $_.ObjectId
            Write-Host $au $user.DisplayName
        }
    }
}

# Start
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')
$tenantName = [System.Environment]::GetEnvironmentVariable('tenantName', 'User')

Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"
Import-Module AzureADPreview
$cred = Get-Creds az-cred.xml
$azAD = Connect-AzureAD -Credential $cred
Import-BulkAzureADAUData "$rootPath\data\teamsInformation.csv"
Add-BulkAzureADAUMemberData "$rootPath\data\users.csv"
Get-BulkAzureADAUMembership
Write-Host "Completed Azure AD Administrative Unit bulk add"
