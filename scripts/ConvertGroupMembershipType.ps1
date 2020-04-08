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
    This script coverts a static group membership to dynamic membership and sets the dynamic rule
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Connect to MicrosoftTeams
$az_cred = Get-Creds az-cred.xml
Connect-AzureAD -Credential $az_cred

# Loads the file groupstoConvert.csv for the groups that must be migrated from static to dynamics teams
Function Read-Groups {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $fileName
    )

    if ($null -eq ($groupsCSV = Read-CSV $fileName)) {
        Write-Host "ERROR - Cannot Process Filename $fileName"
        exit
    }
    else {
        ForEach ($group in $groupsCSV) {
            Set-GroupMembershipType $group.Name $group.DynamicMembershipRule
        }
    }
}

function Set-GroupMembershipType
{
    [CmdletBinding()]
    Param(
        [string]$groupName,
        [string]$dynamicMembershipRule
    )

    $dynamicGroupTypeString = "DynamicMembership"

    If (-not ($group = Get-AzureAdMsGroup -SearchString $groupName)) {
        Write-Host "Group does not exist: $groupName"
    }
    else {
        [System.Collections.ArrayList]$groupTypes = ($group).GroupTypes

        if($null -ne $groupTypes -and $groupTypes.Contains($dynamicGroupTypeString))
        {
            Write-Host "Group is already dynamic membership $teamName"
            Set-Exception -FileName "$rootPath\logs\ConvertGroupMembershipTypeExceptions.csv" -Module "ConvertGroupMembershipType" -Message "Group is already dynamic membership - $groupName"
        }
        else {
            #add the dynamic group type to existing types
            $groupTypes.Add($dynamicGroupTypeString)

            #modify the group properties to make it a static group: i) change GroupTypes to add the dynamic type, ii) start execution of the rule, iii) set the rule
            try {
                Write-Host "changing team membership: "$group.Id $groupTypes
                Set-AzureAdMsGroup -Id $group.Id -GroupTypes $groupTypes -MembershipRuleProcessingState "On" -MembershipRule $dynamicMembershipRule
            }
            catch [Microsoft.Open.MSGraphBeta.Client.ApiException] {
                Write-Host "ERROR ConvertGroupMembershipType - " $_.Exception.ErrorContent.message
                return

            } catch {
                Write-Host "ERROR ConvertGroupMembershipType - convert team membership"
                $Error[0] | Format-List * -Force
                return
            }
        }
    }
}

Write-Host
Write-Host "****************************************************************************************************************" -ForegroundColor Yellow
Write-Host "NOTE!  Changing Azure AD group membership from static to dynamic requires the P1 or above license." -ForegroundColor Yellow
Write-Host "****************************************************************************************************************" -ForegroundColor Yellow
Write-Host
$YesOrNo = Read-Host "DO YOU WANT TO CONTINUE - Please enter your response (y/n)"
while("y","n" -notcontains $YesOrNo )
{
    $YesOrNo = Read-Host "DO YOU WANT TO CONTINUE - Please enter your response (y/n)"
}
if (($YesOrNo).ToUpper() -eq 'n') {
    exit 0
}

Read-Groups -fileName "$rootPath\data\migrateGroups.csv"

