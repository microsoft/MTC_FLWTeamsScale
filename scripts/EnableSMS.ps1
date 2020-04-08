<#
    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED,
    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You
    a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of
    the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which
    the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;
    and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’
    fees, that arise or result from the use or distribution of the Sample Code.
    Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within
    the Premier Customer Services Description.
#>
<#
    Usage:

    NOTE: Dot source the script to load functions into the current session

    . .scripts\EnableSMS.ps1

    It will prompt you for login. (No VS Code Support, passing creds didn't work?)

    Then you can call the following functions:
    Get-UserSMSSignInSettings -userID <userID>
    Set-UserSMSSignInNumber -phoneNumber <phoneNumber> -userID <userID>
    Remove-SMSSignInNumber -userID <userID>
    Enable-UserSMSSignIn -userID <userID>
    Disable-UserSMSSignIn -userID <userID>
    BulkSet-SMSSignInNumber -filePath "<path to file>"
#>

# Get environment variables
$rootPath = [System.Environment]::GetEnvironmentVariable('rootPath', 'User')

# Import the Bulk Add Functions module
Import-Module "$rootPath\scripts\BulkAddFunctions.psm1"

$cred = Get-Creds az-cred.xml

$teamsConnection = Connect-MicrosoftTeams -Credential $cred

$tenant = ""
$clientId = ""
$isPPE=$false

# auth
$authorityUri = if ($isPPE) { "https://login.windows-ppe.net/$tenant" } else { "https://login.microsoftonline.com/$tenant" }
$scopes = if ($isPPE) { 'https://graph.microsoft-ppe.com/.default' } else { 'https://graph.microsoft.com/.default' }
$redirectUri = "urn:ietf:wg:oauth:2.0:oob"

$loginResponse = Get-MsalToken -clientID $clientID -tenantID $tenantDomain -RedirectUri $redirectUri -Authority $authorityUri -Scopes $scopes -Interactive
$accessToken = $loginResponse.AccessToken
$authHeaders = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
$authHeaders.Add('Authorization', 'Bearer ' + $accessToken)
$authHeaders.Add('Content-Type','application/json')
$authHeaders.Add('Accept','application/json, text/plain')
$baseUri = if ($isPPE) { 'https://graph.microsoft-ppe.com/beta/users/' } else { 'https://graph.microsoft.com/beta/users/' }
$phoneAuthMethodUri = "$baseUri{0}/authentication/phoneMethods"
$mobilePhoneMethodId = "3179e48a-750b-4051-897c-87b9720928f7"
$updateMobilePhoneAuthMethodUri = $phoneAuthMethodUri + "/" + $mobilePhoneMethodId
$enableSmsSignInUri = "$updateMobilePhoneAuthMethodUri/enableSmsSignin"
$disableSmsSignInUri = "$updateMobilePhoneAuthMethodUri/disableSmsSignin"
Add-Type -AssemblyName System.Web

function Using-Object
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Object]
        $InputObject,
        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock
    )
    try
    {
        . $ScriptBlock
    }
    finally
    {
        if ($null -ne $InputObject -and $InputObject -is [System.IDisposable])
        {
            $InputObject.Dispose()
        }
    }
}

# Override the built-in cmdlet with a custom version
function Write-Error($message) {
    [Console]::ForegroundColor = 'red'
    [Console]::Error.WriteLine($message)
    [Console]::ResetColor()
}
function HandleHttpErrorResponse($response) {
    if ($response.ContentLength -eq 0) {
        throw New-Object -TypeName System.Web.HttpException -ArgumentList $response.StatusCode
    } else {
        $result = CreateResultFromRespone($response)
        if ($result.GetType().Name -eq "String") {
            return $result;
        } else {
            throw $result;
        }
    }
}
function CreateResultFromRespone($reponse) {
    Using-Object ($reader = New-Object System.IO.StreamReader($response.GetResponseStream())) {
        $reader.BaseStream.Position = 0
        $reader.DiscardBufferedData()
        $responseBody = $reader.ReadToEnd();
        $jsonObj = $responseBody | ConvertFrom-Json
        if ($response.StatusCode -eq 405 -and $jsonObj.error.message -eq "Can not enable Sms sign in on phone auth method as the credential policy is not enabled on user.") {
            return "disabledByPolicy"
        } elseif ($jsonObj.error.code -in @("phoneNumberNotUnique", "disabledByPolicy")) {
            return $jsonObj.error.code
        } else {
            return New-Object -TypeName System.Web.HttpException -ArgumentList $response.StatusCode, $responseBody
        }
    }
}
function Get-UserMobilePhoneMethod
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param(
        [String]$userID
        )
    $uri = $phoneAuthMethodUri -f $userID

    $response = try {
        Invoke-WebRequest -UseBasicParsing -headers $authHeaders -Uri $uri -Method Get
    } catch {
        return HandleHttpErrorResponse($_.Exception.Response)
    }
    $JsonObj = $response.Content | ConvertFrom-Json
    if ($JsonObj.value.count -eq 0) {
        Write-Host "User $userId has no existing mobile phone auth method."
        return $null
    } else {
        $mobilePhoneAuthMethod = $JsonObj.value | Where-Object id -eq $mobilePhoneMethodId | Select-Object -First 1
        if ($mobilePhoneAuthMethod) {
            Write-Host "User $userID's existing mobile phone auth method: " + $mobilePhoneAuthMethod
        } else {
            Write-Host "User $userId has no existing mobile phone auth method."
        }
        return $mobilePhoneAuthMethod;
    }
}
function Get-UserSMSSignInSettings
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param(
        [String]$userID
        )
    $jsonObj = Get-UserMobilePhoneMethod($userID)
    if ($jsonObj) {
        return $jsonObj.smsSignInState
    } else {
        return "notExists"
    }
}
function Enable-UserSMSSignIn
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param(
        [String]$userID
        )
    $jsonObj = Get-UserMobilePhoneMethod($userID)
    if ($jsonObj) {
        $uri = $enableSmsSignInUri -f $userID
        try {
            Invoke-WebRequest -UseBasicParsing -headers $authHeaders -Uri $uri -Method POST | Out-Null
        } catch {
            return HandleHttpErrorResponse($_.Exception.Response)
        }
    } else {
        return "notExists"
    }

    return "success"
}
function Disable-UserSMSSignIn
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param(
        [String]$userID
        )
    $jsonObj = Get-UserMobilePhoneMethod($userID)
    if ($jsonObj) {
        $uri = $disableSmsSignInUri -f $userID
        try {
            Invoke-WebRequest -UseBasicParsing -headers $authHeaders -Uri $uri -Method POST | Out-Null
        } catch {
            return HandleHttpErrorResponse($_.Exception.Response)
        }
    } else {
        return "notExists"
    }

    return "success"
}
function Set-UserSMSSignInNumber
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param(
        [String]$userID,
        [String]$phoneNumber
        )
    $jsonObj = Get-UserMobilePhoneMethod($userID)
    # post/put parameters
    $postParams = @{}
    $postParams.phoneNumber = $phoneNumber
    $postParams.phoneType = "mobile"
    $json = $postparams | ConvertTo-Json -depth 99 -Compress

    # if the user doesn't have a phone number on their account, do a POST request
    # if we want to replace the number with one enable for the policy, do a PUT
    $uri = if ($jsonObj) {
        Write-Host "Update mobile phone auth method to new number..."
        $updateMobilePhoneAuthMethodUri -f $userID
    } else {
        Write-Host "Adding new mobile phone auth method..."
        $phoneAuthMethodUri -f $userID
    }
    $method = if ($jsonObj) {
        "PUT"
    } else {
        "POST"
    }
    $response = try {
        Invoke-WebRequest -UseBasicParsing -headers $authHeaders -Uri $uri -Method $method -Body $json
    } catch {
        return HandleHttpErrorResponse($_.Exception.Response)
    }

    $newJsonObj = $response | ConvertFrom-Json
    Write-Host "New mobile phone auth method: " $newJsonObj
    return $newJsonObj.smsSignInState
}
function Remove-SMSSignInNumber
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param([String]$userID)
    # Get method ID and whether or not this user has a phone number
    $jsonObj = Get-UserMobilePhoneMethod($userID)
    if ($jsonObj) {
        # If there is an SMS Sign-in number, delete it
        Write-Host "Remove existing mobile phone auth method..."
        $uri = $updateMobilePhoneAuthMethodUri -f $userID, $jsonObj.id
        try {
            Invoke-WebRequest -UseBasicParsing -headers $authHeaders -Uri $uri -Method Delete | out-null
        } catch {
            return HandleHttpErrorResponse($_.Exception.Response)
        }
        Return "success"
    } else {
        # If there's no phone number used for SMS Sign-in on the account, return
        return "notExists"
    }
}
function BulkSet-SMSSignInNumber
{
    [cmdletbinding()]
    [Parameter(mandatory=$true)]
    Param([System.IO.FileInfo]$filePath)
    $CSV = Import-Csv $filePath -Header 'userID', 'phoneNumber'
    for ($i = 0; $i -lt $CSV.Count; $i++) {
        $User = $CSV[$i]
        Write-Host "Processing user" $User.userID
        try {
            Set-UserSMSSignInNumber -userID $User.userID -phoneNumber $User.phoneNumber
        } catch {
            Write-Error $_.Exception
        }
    }
}