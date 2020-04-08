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
    This script creates the Teams Channels documented in the CSV file.  This only creates the Teams Channels and their configuration, and it does
    not add users to the Team.
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Connect to MicrosoftTeams
$teams_cred = Get-Creds teams-cred.xml
Connect-MicrosoftTeams -Credential $teams_cred

# For each Team defined in the teamsInformation.csv, create Team Channels defined in the file teamsChannels.csv
# If the Team created in the script CreateTeams.ps1 already existed, that Team will not be updated with the defined channels
Function Read-Channels {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $teamsFile,

        [Parameter(Mandatory)]
        [String]
        $channelsFile
    )

    if ($null -eq ($teamsCSV = Read-CSV $teamsFile)) {
        Write-Host "ERROR CreateTeamsChannel - Cannot Process Filename $fileName"
        exit
    }
    if ($null -eq ($channelsCSV = Read-CSV $channelsFile)) {
        Write-Host "ERROR CreateTeamsChannel - Cannot Process Filename $fileName"
        exit
    }

    ForEach ($team in $teamsCSV) {
        try {
            if (-not($existingTeam = Get-Team -DisplayName $team.teamName -ErrorAction SilentlyContinue)) {
                Write-Host "Team not found: $team"
            }
            else {
                ForEach ($channel in $channelsCSV) {
                    Add-Channel $existingTeam.DisplayName $existingTeam.GroupId $channel.channelName $channel.channelDescription $channel.membershipType $channel.privateChannelOwner
                }
            }
        }
        catch {
            Write-Host "ERROR CreateTeamsChannels - get team"
            $Error[0] | Format-List * -Force
        }
    }
}

# Test if the Team already exists, and if not create the Team according to the parameters
Function Add-Channel {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $teamDisplayName,

        [Parameter(Mandatory)]
        [String]
        $teamGroupId,

        [Parameter(Mandatory)]
        [String]
        $channelName,

        [Parameter(Mandatory)]
        [String]
        $channelDescription,

        [Parameter(Mandatory)]
        [String]
        $membershipType,

        [Parameter()]
        [String]
        $privateChannelOwner
    )

    try {
        Write-Host "CREATING CHANNEL for TEAM: Channel $channelName, Team $teamDisplayName"
        New-TeamChannel -GroupId $teamGroupId -DisplayName $channelName -Description $channelDescription -MembershipType $membershipType -ErrorAction SilentlyContinue
    }
    catch {
        Write-Host "ERROR CreateTeamsChannels - creating team channel"
        $Error[0] | Format-List * -Force
        Set-Exception -FileName "$rootPath\logs\CreateTeamsChannelsExceptions.csv" -Module "CreateTeamsChannels" -Message "Teams Channel Already Exists $Error[0]"
    }
}
Read-Channels "$rootPath\data\teamsInformation.csv" "$rootPath\data\teamsChannels.csv"

Write-Host "Completed Teams Channels"
