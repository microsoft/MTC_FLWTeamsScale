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
        This script creates the users documented in the firstline worker and firestline manager CSV files.  Two security groups are also created
        named "Firstline Workers" and "Firstline Managers".  The workers and managers are added to their respective security groups.
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')
$tenantName = [System.Environment]::GetEnvironmentVariable('tenantName', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Connect to AzureAD
$azCred = Get-Creds az-cred.xml
Connect-AzureAD -Credential $azCred

# Create CsOnlineSession
$csolCred = Get-Creds csol-cred.xml
$sfbSession = New-CsOnlineSession -Credential $csolCred
Import-PSSession $sfbSession

# Connect to MicrosoftTeams
$teamsCred = Get-Creds teams-cred.xml
Connect-MicrosoftTeams -Credential $teamsCred

# Connect to MSOLService
$msolCred = Get-Creds msol-cred.xml
Connect-MSOLService -Credential $msolCred

# Create Azure AD Security Groups
Write-Host "Creating security groups"
if ($null -eq ($secgrpCSV = Read-CSV "$rootPath\data\securityGroups.csv")) {
    Write-Host "ERROR - Cannot Process Filename $fileName"
    exit
}
else {
    ForEach ($secGrp in $secgrpCSV) {
        try {
            if (-not(Get-AzureADGroup -SearchString $secGrp.name)) {
                New-AzureADGroup -DisplayName $secGrp.Name -Description $secGrp.Description -MailEnabled $false -SecurityEnabled $true -MailNickName "NotSet"
            }
            else {
                Write-Host "Security Group already exists: " $secGrp.name
                $name = $secGrp.Name
                Set-Exception -FileName "$rootPath\logs\SecurityGroupsExceptions.csv" -Module "CreateUsers" -Message "Security Group already exists: $name"
            }
        }
        catch {
            Write-Host "ERROR - creating security group "
            $Error[0] | Format-List * -Force
        }
    }
}

# Add the users and assign to their security group

Write-Host "Creating users and assigning them to security group"
if ($null -eq ($usersCSV = Read-CSV "$rootPath\data\users.csv")) {
    Write-Host "ERROR - Cannot Process Filename $fileName"
    exit
}
else {
    ForEach ($user in $usersCSV) {
        if (-not(Get-AzureADUser -SearchString $user.name)) {
            $upn = $user.name + "@" + $tenantName
            try {
                $passwordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
                $passwordProfile.Password = "TempPwd123"
                $passwordProfile.ForceChangePasswordNextLogin = $true

                $userObj = New-AzureADUser -DisplayName $user.name -City $user.City -State $user.State -Country $user.Region -GivenName $user.name -JobTitle $user.JobTitle  -MailNickName $user.name `
                    -UserPrincipalName $upn -AccountEnabled $true -PasswordProfile $passwordProfile -UsageLocation $user.UsageLocation

                # Add user to security group
                $secGroup = Get-AzureADGroup -SearchString $user.SecurityGroup
                Add-AzureADGroupMember -ObjectId $secGroup.Objectid -RefObjectId $userObj.ObjectId
            }
            catch {
                Write-Host "ERROR - creating firstline user"
                $Error[0] | Format-List * -Force
            }
        }
        else {
            Write-Host "User already exists: " $user.Name
            $name = $user.Name
            Set-Exception -FileName "$rootPath\logs\CreateUsersExceptions.csv" -Module "CreateUsers" -Message "User Already Exists $name"
        }
    }
}

# Assign licensing to the users according to the licensing plan set for their security group
Write-Host "Assigning licenses"
$license = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
$assignedLicense = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses

ForEach ($secGrp in $secgrpCSV) {
    $group = Get-AzureADGroup -SearchString $secGrp.Name
    $groupMembers = Get-AzureADGroupMember -ObjectId $group.ObjectId -All $true | Where-Object { $_.ObjectType -eq "User" }
    try {
        ForEach ($member in $groupMembers) {
            $license.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $secGrp.LicensePlan -EQ).SkuID
            $assignedLicense.AddLicenses = $license
            Set-AzureADUserLicense -ObjectId $member.ObjectId -AssignedLicenses $assignedLicense
        }
    }
    catch {
        Write-Host "ERROR - creating firstline user"
        $Error[0] | Format-List * -Force
    }
}
Write-Host "Completed Create Users"
Remove-PSSession -Session $sfbSession
