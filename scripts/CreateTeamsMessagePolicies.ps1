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
    This script creates the Teams Message policies
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

# Get CSOnline session credentials
$teams_cred = Get-Creds teams-cred.xml
$sfbSession = New-CsOnlineSession -Credential $teams_cred
Import-PSSession $sfbSession

Function Read-MessagingPolicies {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $fileName
    )

    if ($null -eq ($policyCSV = Read-CSV $fileName)) {
        Write-Host "ERROR - Cannot Process Filename $fileName"
        exit
    }
    else {
        ForEach ($policy in $policyCSV) {
            # this approach allows the CSV to contain the text true/false for boolean parameters and it converts the param to a bool type to pass to the function
            $allowUrlPreview = [System.Convert]::ToBoolean($policy.allowUrlPreview)
            $allowOwnerDeleteMessage = [System.Convert]::ToBoolean($policy.allowOwnerDeleteMessage)
            $allowUserEditMessage = [System.Convert]::ToBoolean($policy.allowUserEditMessage)
            $allowUserDeleteMessage = [System.Convert]::ToBoolean($policy.allowUserDeleteMessage)
            $allowUserChat = [System.Convert]::ToBoolean($policy.allowUserChat)
            $allowRemoveUser = [System.Convert]::ToBoolean($policy.allowRemoveUser)
            $allowGiphy = [System.Convert]::ToBoolean($policy.allowGiphy)
            $allowMemes = [System.Convert]::ToBoolean($policy.allowMemes)
            $allowImmersiveReader = [System.Convert]::ToBoolean($policy.allowImmersiveReader)
            $allowStickers = [System.Convert]::ToBoolean($policy.allowStickers)
            $allowUserTranslation = [System.Convert]::ToBoolean($policy.allowUserTranslation)
            $allowPriorityMessages = [System.Convert]::ToBoolean($policy.allowPriorityMessages)

            Add-MessagingPolicy $policy.policyName $policy.policyDescription $allowUrlPreview  $allowOwnerDeleteMessage $allowUserEditMessage `
                $allowUserDeleteMessage $allowUserChat $allowRemoveUser $allowGiphy $policy.giphyRatingType $allowMemes `
                $allowImmersiveReader $allowStickers $allowUserTranslation $policy.readReceiptsEnabledType $allowPriorityMessages `
                $policy.channelsInChatListEnabledType $policy.audioMessageEnabledType
        }
    }
}

Function Add-MessagingPolicy {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $policyName,

        [Parameter(Mandatory)]
        [String]
        $policyDescription,

        [Parameter(Mandatory)]
        [Bool]
        $allowUrlPreview,

        [Parameter(Mandatory)]
        [Bool]
        $allowOwnerDeleteMessage,

        [Parameter(Mandatory)]
        [Bool]
        $allowUserEditMessage,

        [Parameter(Mandatory)]
        [Bool]
        $allowUserDeleteMessage,

        [Parameter(Mandatory)]
        [Bool]
        $allowUserChat,

        [Parameter(Mandatory)]
        [Bool]
        $allowRemoveUser,

        [Parameter(Mandatory)]
        [Bool]
        $allowGiphy,

        [Parameter(Mandatory)]
        [String]
        $giphyRatingType,

        [Parameter(Mandatory)]
        [Bool]
        $allowMemes,

        [Parameter(Mandatory)]
        [Bool]
        $allowImmersiveReader,

        [Parameter(Mandatory)]
        [Bool]
        $allowStickers,

        [Parameter(Mandatory)]
        [Bool]
        $allowUserTranslation,

        [Parameter(Mandatory)]
        [String]
        $readReceiptsEnabledType,

        [Parameter(Mandatory)]
        [Bool]
        $allowPriorityMessages,

        [Parameter(Mandatory)]
        [String]
        $channelsInChatListEnabledType,

        [Parameter(Mandatory)]
        [String]
        $audioMessageEnabledType
    )

    try {
        if (-not(New-CsTeamsMessagingPolicy -Identity $policyName -Description $policyDescription -AllowUrlPreview $allowUrlPreview `
                    -AllowOwnerDeleteMessage $allowOwnerDeleteMessage -AllowUserEditMessage $allowUserEditMessage -AllowUserDeleteMessage $allowUserDeleteMessage `
                    -AllowUserChat $allowUserChat -AllowRemoveUser $allowRemoveUser -AllowGiphy $allowGiphy -GiphyRatingType $giphyRatingType -AllowMemes $allowMemes `
                    -AllowImmersiveReader $allowImmersiveReader -AllowStickers $allowStickers -AllowUserTranslation $allowUserTranslation `
                    -ReadReceiptsEnabledType $readReceiptsEnabledType -AllowPriorityMessages $allowPriorityMessages `
                    -ChannelsInChatListEnabledType $channelsInChatListEnabledType -AudioMessageEnabledType $audioMessageEnabledType -ErrorAction SilentlyContinue)) {
            Write-Host "Messaging Policy Already Exists $policyName"
            Set-Exception -FileName "$rootPath\logs\CreateTeamsMessagePoliciesExceptions.csv" -Module "CreateTeamsMessagePolicies" -Message "Messaging Policy Already Exists $policyName"
        }
        else {
            Write-Host "Policy created: $policyName"
        }
    }
    catch {
        Write-Host "ERROR CreateTeamsMessagePolicies - creating policy"
        $Error[0] | Format-List * -Force
    }
}

Read-MessagingPolicies "$rootPath\data\teamsMessagePolicies.csv"
Write-Host "Completed Teams Messaging Policy"
Remove-PSSession -Session $sfbSession




