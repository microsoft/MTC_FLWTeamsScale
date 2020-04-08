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
    This script creates the Teams documented in the CSV file.  This only creates the Teams and their default configuration, and it does
    not add users to the Team.
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Connect to MicrosoftTeams
$teams_cred = Get-Creds teams-cred.xml
Connect-MicrosoftTeams -Credential $teams_cred

# Loads the file teamsInformation.csv for the Teams needing to be created and calls to create each Team
Function Read-Teams {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $fileName
    )

    if ($null -eq ($teamsCSV = Read-CSV $fileName)) {
        Write-Host "ERROR - Cannot Process Filename $fileName"
        exit
    }
    else {
        ForEach ($team in $teamsCSV) {
            Add-Team $team.region $team.district $team.location $team.storeNumber $team.TeamName $team.TeamDescription $team.IsTeamPrivate  $team.TeamOwners $team.TeamEmailNickname
        }
    }
}

# Test if the Team already exists, and if not create the Team according to the parameters
Function Add-Team {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $region,

        [Parameter(Mandatory)]
        [String]
        $district,

        [Parameter(Mandatory)]
        [String]
        $location,

        [Parameter(Mandatory)]
        [String]
        $storeNumber,

        [Parameter(Mandatory)]
        [String]
        $teamName,

        [Parameter(Mandatory)]
        [String]
        $teamDescription,

        [Parameter(Mandatory)]
        [String]
        $isTeamPrivate,

        [Parameter(Mandatory)]
        [String]
        $teamOwners,

        [Parameter(Mandatory)]
        [String]
        $teamEmailNickname
    )

    if (-not(Get-Team -DisplayName $teamName -ErrorAction SilentlyContinue)) {
        try {
            Write-Host "CREATING TEAM: Name $teamName"
            New-Team -DisplayName $teamName -Visibility $isTeamPrivate -Description $teamDescription -Owner $teamOwners -MailNickName $teamEmailNickname -ErrorAction SilentlyContinue
        }
        catch {
            Write-Host "ERROR CreateTeams- creating Team"
            $Error[0] | Format-List * -Force
        }
    }
    else {
        Write-Host "Team Already Exists $teamName"
        Set-Exception -FileName "$rootPath\logs\CreateTeamsExceptions.csv" -Module "CreateTeams" -Message "Team Already Exists $teamName"
    }
}
Read-Teams "$rootPath\data\teamsInformation.csv"

Write-Host "Completed Create Teams"