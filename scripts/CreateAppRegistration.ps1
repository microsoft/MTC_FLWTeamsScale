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

# Import the Bulk Add Functions module
Import-Module .\scripts\BulkAddFunctions.psm1

# Connect to Microsoft Teams
$az_cred = Get-Creds az-cred.xml
Connect-MicrosoftTeams -Credential $az_cred

$appName = "TeamsBulkAdd"
$appURI = "https://localhost"
$appReplyURLs = @($appURI)

Write-Host "Getting new App Reg $appName"
if (!($myApp = Get-AzureADApplication -Filter "DisplayName eq '$($appName)'"  -ErrorAction SilentlyContinue)) {
    Write-Host "Creating Azure AD App Registration $appName" -ForegroundColor Yellow
    $myApp = New-AzureADApplication -DisplayName $appName -IdentifierUris $appURI -Homepage $appHomePageUrl -ReplyUrls $appReplyURLs
}

Write-Host "Getting service principal"

$resourceAccess = New-Object -TypeName "Microsoft.Open.AzureAD.Model.RequiredResourceAccess"
$res1 = New-Object -TypeName "microsoft.open.azuread.model.resourceAccess" -ArgumentList "40e263e50-5827-48a4-b97c-d940288653c7", "scope"

$resourceAccess.ResourceAccess = $res1
$resourceAccess.ResourceAppId = "123"

$resourceAccess

try {
    Set-AzureADApplication -ObjectId $myApp.ObjectId -RequiredResourceAccess $resourceAccess -ErrorAction stop
}
catch {
    Write-Host "ERROR Create App Registration - set Azure AD application"
    $Error[0] | Format-List * -Force
}


