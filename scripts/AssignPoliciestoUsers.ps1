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
        This script assigns Teams policies to the user
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# maxPolicyBatchSize is the current limit for calling batch policy assignments.  This number should only change if the max limit per batch is increased
$maxPolicyBatchSize = 5000

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

# Load the security groups CSV
if ($null -eq ($secgrpCSV = Read-CSV "$rootPath\data\securityGroups.csv")) {
    Write-Host "ERROR - Cannot Process Filename $fileName"
    exit
}

# Assign Teams policies to the users
Write-Host "Assigning Messaging, App Setup Policy, and App Permission Policy"
ForEach ($secGrp in $secgrpCSV) {
    try {
        $group = Get-AzureADGroup -SearchString $secGrp.Name
        $groupMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true | Where-Object { $_.ObjectType -eq "User" }

        # conform to the current limit for calling batch policy assignments
        $numBatches = Set-BatchIterations $groupMembers.Count $maxPolicyBatchSize

        # Each of these loops are used to manage batching new policy assignments to the $maxPolicyBatchSize set at the beginning of this script.  This is a throttling control that Microsoft
        # sets and should only be changed if Microsoft changes the limit
        $batchStart = 0
        For ($i = 1; $i -le $numBatches; $i++) {
            Write-Host "Assigning Teams Messaging Policy to Security Group  (batch $i of $numBatches)  " $secGrp.name
            $upperLimit = Set-NextBatch -BatchStart $batchStart -MaxBatchSize $maxPolicyBatchSize -UserCount $groupMembers.Count
            $memberBatchArray = [System.Collections.ArrayList]@()
            For ($j = $batchStart; $j -lt $upperLimit; $j++){
                $null = $memberBatchArray.Add($groupMembers[$j])
            }
            New-CsBatchPolicyAssignmentOperation -PolicyType TeamsMessagingPolicy -PolicyName $secGrp.MessagePolicy -Identity $memberBatchArray.UserPrincipalName -OperationName "Batch assign Teams message policy"
            $batchStart = $upperLimit
        }

        $batchStart = 0
        For ($i = 1; $i -le $numBatches; $i++) {
            Write-Host "Assigning Teams App Setup Policy to Security Group (batch $i of $numBatches) "  $secGrp.name
            $upperLimit = Set-NextBatch -BatchStart $batchStart -MaxBatchSize $maxPolicyBatchSize -UserCount $groupMembers.Count
            $null = $memberBatchArray = [System.Collections.ArrayList]@()
            For ($j = $batchStart; $j -lt $upperLimit; $j++){
                $null = $memberBatchArray.Add($groupMembers[$j])
            }
            New-CsBatchPolicyAssignmentOperation -PolicyType TeamsAppSetupPolicy -PolicyName $secGrp.AppSetupPolicy -Identity $memberBatchArray.UserPrincipalName -OperationName "Batch assign Teams app setup policy"
            $batchStart = $upperLimit
        }

        $batchStart = 0
        For ($i = 1; $i -le $numBatches; $i++) {
            Write-Host "Assigning Teams App Permission Policy to Security Group  (batch $i of $numBatches)  "  $secGrp.name
            $upperLimit = Set-NextBatch -BatchStart $batchStart -MaxBatchSize $maxPolicyBatchSize -UserCount $groupMembers.Count
            $memberBatchArray = [System.Collections.ArrayList]@()
            For ($j = $batchStart; $j -lt $upperLimit; $j++){
                $null = $memberBatchArray.Add($groupMembers[$j])
            }
            New-CsBatchPolicyAssignmentOperation -PolicyType TeamsAppPermissionPolicy -PolicyName $secGrp.AppPermissionsPolicy -Identity $memberBatchArray.UserPrincipalName -OperationName "Batch assign Teams app permission policy"
            $batchStart = $upperLimit
        }
    }
    catch {
        Write-Host "ERROR AssignPoliciestoUsers - Assigning Teams Policies"
        $Error[0] | Format-List * -Force
    }
}
