<#
        This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
        THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
        INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You
        a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of
        the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which
        the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;
        and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneysÃ¢â‚¬â„¢
        fees, that arise or result from the use or distribution of the Sample Code.
        Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within
        the Premier Customer Services Description.
#>

<#
        This script assigns users to their assigned Team
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')
$tenantName = [System.Environment]::GetEnvironmentVariable('tenantName', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Connect to AzureAD
$azCred = Get-Creds az-cred.xml
Connect-AzureAD -Credential $azCred

# Connect to MicrosoftTeams
$teamsCred = Get-Creds teams-cred.xml
Connect-MicrosoftTeams -Credential $teamsCred

# Connect to MSOLService
$msolCred = Get-Creds msol-cred.xml
Connect-MSOLService -Credential $msolCred

#build an array of Team names and object IDs
$teamArr = @()
if ($null -eq ($teamsCSV = Read-CSV "$rootPath\data\teamsInformation.csv")) {
    Write-Host "ERROR - Cannot Process Filename $fileName"
    exit
}
else {
    ForEach ($team in $teamsCSV) {
        try {
            $teamArr += Get-Team | Where-Object { $_.DisplayName -eq $team.TeamName } | Select-Object GroupId, DisplayName
        }
        catch {
            Write-Host "ERROR AssignUserstoTeams - building Teams array"
            $Error[0] | Format-List * -Force
        }
    }
}

Write-Host "Assiging users to their Teams"
if ($null -eq ($usersCSV = Read-CSV "$rootPath\data\users.csv")) {
    Write-Host "ERROR - Cannot Process Filename $fileName"
    return
}
else {
    ForEach ($user in $usersCSV) {
        if (Get-AzureADUser -SearchString $user.name) {
            $upn = $user.name + "@" + $tenantName
            try {
                $objId = $teamArr | Where-Object { $_.DisplayName -eq $user.TeamName } | Select-Object -ExpandProperty GroupId
                Add-TeamUser -GroupId $objId -User $upn
            }
            catch {
                Write-Host "ERROR AssignUserstoTeams - assigning user to Team"
                $Error[0] | Format-List * -Force
            }
        }
        else {
            Write-Host "User not found assigning to Team: " $user.name
        }
    }
}

Write-Host "Completed assigning users to teams"
