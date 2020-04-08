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
    This module contains commonly used helper Functions
#>

# Helper function to read CSV files
Function Read-CSV {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $fileName
    )

    If (-not($array = Get-Content $fileName -ErrorAction SilentlyContinue | ConvertFrom-Csv)) {
        Write-Host "ERROR FILE NOT FOUND: $fileName"
        return $null
    }
    return $array
}

# Helper function to return the credentials for a service endpoint
Function Get-Creds {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]
        $FileName
    )
    $global:az_keypath = "$ENV:LOCALAPPDATA\keys"
    $appreg_cred = Import-Clixml -Path $global:az_keyPath\$FileName

    return $appreg_cred
}

# Helper function to set the credentials for a service endpoint
Function Set-Creds {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String]
        $FileName
    )

    $global:az_keypath = "$ENV:LOCALAPPDATA\keys"

    if (!(Test-Path $global:az_keypath)) {
        mkdir $global:az_keypath
    }

    $appreg_cred = Get-Credential
    $appreg_cred | Export-Clixml -Path $global:az_keyPath\$FileName -Confirm

    return $appreg_cred
}

# Helper function to output exceptions in processing, i.e. where something already exists
Function Set-Exception {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [String]
        $FileName,

        [Parameter(Mandatory)]
        [String]
        $Module,

        [Parameter(Mandatory)]
        [String]
        $Message
    )
    $EvtProperties = @{
        Module   = $Module
        Message  = $Message
    }
    $PSObj = New-Object -TypeName psobject -Property $EvtProperties
    $PSObj | export-csv $fileName -Append
}

# Helper function to determine number of batches required for large scale, batch operations
Function Set-BatchIterations {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Int32]
        $userCount,

        [Parameter(Mandatory)]
        [Int32]
        $batchSize
    )
    if ($userCount -gt $batchSize) {
        $batchIterations = [math]::Round($userCount / $batchSize)
        if (($userCount % $batchSize) -gt 0) {
            $batchIterations += 1
        }
    }
    else {
        $batchIterations = 1
    }
    return $batchIterations
}

# Helper function to determine the upper limit of the next processing batch required for large scale, batch operations
Function Set-NextBatch {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Int32]
        $BatchStart,

        [Parameter(Mandatory)]
        [Int32]
        $MaxBatchSize,

        [Parameter(Mandatory)]
        [Int32]
        $UserCount
    )
    $UpperLimit = $BatchStart
    if (($UpperLimit + $MaxBatchSize) -le $UserCount) {
        $UpperLimit += $MaxBatchSize
    }
    else {
        $UpperLimit = $UserCount
    }
    return $UpperLimit
}

