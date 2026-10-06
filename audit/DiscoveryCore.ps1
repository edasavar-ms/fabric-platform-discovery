# Internal helpers and the original 32-column workspace renderer.
# Loaded by FabricAudit.psm1; this file has no standalone entry point.
# Prepared by Eda Avar, Solution Engineer, Microsoft, as presales discovery input.
# See DISCLAIMER.md: unsupported sample code, provided as is.
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

$PowerBIBaseUri = "https://api.powerbi.com/v1.0/myorg"
$FabricBaseUri = "https://api.fabric.microsoft.com/v1"
$LargeModelThresholdBytes = 1GB

# Power BI-only types are excluded from FabricItemsCount. Dataflow is not
# excluded because the Fabric endpoint represents Dataflow Gen2, while the
# legacy DataflowsCount field represents Power BI dataflows. A deny-list also
# ensures newly introduced Fabric item types are included automatically.
$NonFabricItemTypes = @(
    "Dashboard",
    "Report",
    "SemanticModel",
    "PaginatedReport",
    "Datamart",
    "App"
)

$script:TokenCache = @{}

function Get-PropertyValue {
    <#
    Reads an optional API property without failing under strict mode. Fabric and
    Power BI omit properties that do not apply to a particular object.
    #>
    param(
        [Parameter()]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter()]$DefaultValue = $null
    )

    if ($null -eq $InputObject) {
        return $DefaultValue
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }

    return $property.Value
}

function ConvertFrom-TokenValue {
    <#
    Az.Accounts returns tokens as SecureString in current releases and as plain
    strings in some older releases. This function supports both representations
    without writing the token to the console or disk.
    #>
    param([Parameter(Mandatory = $true)]$Token)

    if ($Token -is [System.Security.SecureString]) {
        $credential = New-Object System.Management.Automation.PSCredential(
            "token",
            $Token
        )
        return $credential.GetNetworkCredential().Password
    }

    return [string]$Token
}

function Get-ApiAuthorizationHeader {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("PowerBI", "Fabric")]
        [string]$Audience,

        [Parameter()]
        [switch]$ForceRefresh
    )

    $resourceUrl = if ($Audience -eq "PowerBI") {
        "https://analysis.windows.net/powerbi/api"
    }
    else {
        "https://api.fabric.microsoft.com"
    }

    $cachedToken = $script:TokenCache[$Audience]
    $refreshBefore = [DateTimeOffset]::UtcNow.AddMinutes(5)
    if (
        -not $ForceRefresh -and
        $null -ne $cachedToken -and
        $cachedToken.ExpiresOn -gt $refreshBefore
    ) {
        return @{ Authorization = "Bearer $($cachedToken.AccessToken)" }
    }

    $tokenResult = Get-AzAccessToken -ResourceUrl $resourceUrl -ErrorAction Stop
    $accessToken = ConvertFrom-TokenValue -Token $tokenResult.Token
    $expiresOn = [DateTimeOffset]$tokenResult.ExpiresOn

    $script:TokenCache[$Audience] = @{
        AccessToken = $accessToken
        ExpiresOn   = $expiresOn
    }

    return @{ Authorization = "Bearer $accessToken" }
}

function Get-HttpStatusCode {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    $response = Get-PropertyValue `
        -InputObject $ErrorRecord.Exception `
        -Name "Response"
    if ($null -eq $response) {
        return $null
    }

    try {
        return [int]$response.StatusCode
    }
    catch {
        try {
            return [int]$response.StatusCode.value__
        }
        catch {
            return $null
        }
    }
}

function Get-RetryAfterSeconds {
    param(
        [Parameter(Mandatory = $true)]$ErrorRecord,
        [Parameter(Mandatory = $true)][int]$Attempt
    )

    $response = Get-PropertyValue `
        -InputObject $ErrorRecord.Exception `
        -Name "Response"
    if ($null -ne $response) {
        try {
            $headerValue = $response.Headers["Retry-After"]
            if ($null -ne $headerValue) {
                return [Math]::Max(1, [int]$headerValue)
            }
        }
        catch {
            # The response type differs between Windows PowerShell and
            # PowerShell 7. Fall through to exponential backoff when the header
            # cannot be read through the common indexer.
        }

        try {
            if ($null -ne $response.Headers.RetryAfter.Delta) {
                return [Math]::Max(
                    1,
                    [int][Math]::Ceiling(
                        $response.Headers.RetryAfter.Delta.TotalSeconds
                    )
                )
            }
        }
        catch {
            # Try the HTTP-date representation below.
        }

        try {
            if ($null -ne $response.Headers.RetryAfter.Date) {
                $delay = (
                    $response.Headers.RetryAfter.Date -
                    [DateTimeOffset]::UtcNow
                ).TotalSeconds
                return [Math]::Max(
                    1,
                    [int][Math]::Ceiling($delay)
                )
            }
        }
        catch {
            # Use exponential backoff below.
        }
    }

    return [Math]::Min(60, [Math]::Pow(2, $Attempt))
}

function Get-ScannedWorkspaceDetails {
    <#
    Scanner requests accept up to 100 workspace IDs. Each batch is started,
    polled to a terminal state, and then read. A failed batch is recorded
    against its workspaces so the CSV can be exported as Partial without
    presenting missing RBAC data as a successful orphan result.
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Workspaces,
        [Parameter()][ValidateRange(1, 100)][int]$BatchSize = $ScanBatchSize
    )

    $detailsById = @{}
    $errorsById = @{}

    for ($offset = 0; $offset -lt $Workspaces.Count; $offset += $BatchSize) {
        $lastIndex = [Math]::Min(
            $offset + $BatchSize - 1,
            $Workspaces.Count - 1
        )
        $batch = @($Workspaces[$offset..$lastIndex])
        $workspaceIds = @($batch | ForEach-Object { [string]$_.id })
        $batchNumber = [int][Math]::Floor($offset / $BatchSize) + 1
        $batchCount = [int][Math]::Ceiling(
            $Workspaces.Count / [double]$BatchSize
        )

        Write-Host (
            "Scanning workspace batch {0} of {1} ({2} workspace(s))..." -f
            $batchNumber,
            $batchCount,
            $workspaceIds.Count
        )

        try {
            $scanRequest = Invoke-DiscoveryApiRequest `
                -Uri "$PowerBIBaseUri/admin/workspaces/getInfo?getArtifactUsers=true" `
                -Audience PowerBI `
                -Method Post `
                -Body @{ workspaces = $workspaceIds } `
                -Operation "Start workspace scan batch $batchNumber"

            $scanId = [string](Get-PropertyValue `
                -InputObject $scanRequest `
                -Name "id" `
                -DefaultValue "")
            if ([string]::IsNullOrWhiteSpace($scanId)) {
                throw "Workspace scan batch $batchNumber returned no scan ID."
            }

            $deadline = [DateTimeOffset]::UtcNow.AddSeconds(
                $ScanTimeoutSeconds
            )
            do {
                Wait-AuditSeconds 2
                $statusResponse = Invoke-DiscoveryApiRequest `
                    -Uri "$PowerBIBaseUri/admin/workspaces/scanStatus/$scanId" `
                    -Audience PowerBI `
                    -Operation "Get status for workspace scan $scanId"
                $scanStatus = [string](Get-PropertyValue `
                    -InputObject $statusResponse `
                    -Name "status" `
                    -DefaultValue "")

                if ($scanStatus -eq "Failed") {
                    $scanError = Get-PropertyValue `
                        -InputObject $statusResponse `
                        -Name "error"
                    $scanErrorMessage = [string](Get-PropertyValue `
                        -InputObject $scanError `
                        -Name "message" `
                        -DefaultValue "No error message was returned.")
                    throw (
                        "Workspace scan $scanId failed. $scanErrorMessage"
                    )
                }

                if ([DateTimeOffset]::UtcNow -gt $deadline) {
                    throw (
                        "Workspace scan $scanId did not complete within " +
                        "$ScanTimeoutSeconds seconds."
                    )
                }
            } while ($scanStatus -ne "Succeeded")

            $scanResult = Invoke-DiscoveryApiRequest `
                -Uri "$PowerBIBaseUri/admin/workspaces/scanResult/$scanId" `
                -Audience PowerBI `
                -Operation "Get result for workspace scan $scanId"
            $scannedWorkspaces = @(Get-PropertyValue `
                -InputObject $scanResult `
                -Name "workspaces" `
                -DefaultValue @())

            foreach ($scannedWorkspace in $scannedWorkspaces) {
                $scannedId = [string](Get-PropertyValue `
                    -InputObject $scannedWorkspace `
                    -Name "id" `
                    -DefaultValue "")
                if (-not [string]::IsNullOrWhiteSpace($scannedId)) {
                    $detailsById[$scannedId] = $scannedWorkspace

                    $retrievalState = [string](Get-PropertyValue `
                        -InputObject $scannedWorkspace `
                        -Name "dataRetrievalState" `
                        -DefaultValue "")
                    if ($retrievalState -match "(?i)fail|error|partial") {
                        $errorsById[$scannedId] = (
                            "Workspace scanner data retrieval state was " +
                            "'$retrievalState'."
                        )
                    }

                    $scannedType = [string](Get-PropertyValue `
                        -InputObject $scannedWorkspace `
                        -Name "type" `
                        -DefaultValue "")
                    $usersProperty = $scannedWorkspace.PSObject.Properties[
                        "users"
                    ]
                    if (
                        $scannedType -in @("Workspace", "Group") -and
                        $null -eq $usersProperty
                    ) {
                        $errorsById[$scannedId] = (
                            "The scanner did not return the requested " +
                            "workspace users property."
                        )
                    }
                }
            }

            foreach ($workspaceId in $workspaceIds) {
                if (-not $detailsById.ContainsKey($workspaceId)) {
                    $errorsById[$workspaceId] = (
                        "The completed scanner result did not contain this " +
                        "workspace."
                    )
                }
            }
        }
        catch {
            $message = $_.Exception.Message
            if ($workspaceIds.Count -gt 1) {
                Write-Warning (
                    "Workspace scan batch $batchNumber failed. Retrying each " +
                    "workspace separately to isolate the failure. $message"
                )
                foreach ($singleWorkspace in $batch) {
                    $singleResult = Get-ScannedWorkspaceDetails `
                        -Workspaces @($singleWorkspace) `
                        -BatchSize 1
                    foreach ($key in $singleResult.DetailsById.Keys) {
                        $detailsById[$key] = $singleResult.DetailsById[$key]
                    }
                    foreach ($key in $singleResult.ErrorsById.Keys) {
                        $errorsById[$key] = $singleResult.ErrorsById[$key]
                    }
                }
            }
            else {
                Write-Warning (
                    "Workspace scan batch $batchNumber failed. $message"
                )
                $errorsById[$workspaceIds[0]] = $message
            }
        }
    }

    return [PSCustomObject]@{
        DetailsById = $detailsById
        ErrorsById  = $errorsById
    }
}

function Get-WorkspaceHostingMode {
    param(
        [Parameter(Mandatory = $true)]$Workspace,
        [Parameter()]$Capacity
    )

    $isDedicated = [bool](Get-PropertyValue `
        -InputObject $Workspace `
        -Name "isOnDedicatedCapacity" `
        -DefaultValue $false)

    if (-not $isDedicated) {
        return "Shared capacity"
    }

    $sku = [string](Get-PropertyValue `
        -InputObject $Capacity `
        -Name "sku" `
        -DefaultValue "")
    if ($sku -match "^F\d+") {
        return "Fabric capacity"
    }
    if ($sku -match "^P\d+") {
        return "Power BI Premium capacity"
    }
    if ($sku -match "^(A|EM)\d+") {
        return "Power BI Embedded capacity"
    }
    if ($sku -match "^PPU") {
        return "Premium Per User"
    }

    return "Dedicated capacity"
}

function Get-PrincipalTypeLabel {
    param([Parameter()]$Principal)

    $profile = Get-PropertyValue -InputObject $Principal -Name "profile"
    if ($null -ne $profile) {
        return "Service Principal Profile"
    }

    $principalType = [string](Get-PropertyValue `
        -InputObject $Principal `
        -Name "principalType" `
        -DefaultValue "")

    switch ($principalType.ToLowerInvariant()) {
        "user" { return "Individual" }
        "group" { return "AD Group" }
        "app" { return "Service Principal" }
        "serviceprincipal" { return "Service Principal" }
        "serviceprincipalprofile" { return "Service Principal Profile" }
        "none" { return "Entire Tenant" }
        default {
            if ([string]::IsNullOrWhiteSpace($principalType)) {
                return ""
            }
            return $principalType
        }
    }
}

function Get-PrincipalLabel {
    param([Parameter(Mandatory = $true)]$Principal)

    $profile = Get-PropertyValue -InputObject $Principal -Name "profile"
    $identityObject = if ($null -ne $profile) { $profile } else { $Principal }
    $displayName = [string](Get-PropertyValue `
        -InputObject $identityObject `
        -Name "displayName" `
        -DefaultValue "")
    $identifierProperty = if ($null -ne $profile) { "id" } else { "identifier" }
    $identifier = [string](Get-PropertyValue `
        -InputObject $identityObject `
        -Name $identifierProperty `
        -DefaultValue "")

    if (
        -not [string]::IsNullOrWhiteSpace($displayName) -and
        -not [string]::IsNullOrWhiteSpace($identifier) -and
        $displayName -ne $identifier
    ) {
        return "$displayName <$identifier>"
    }
    if (-not [string]::IsNullOrWhiteSpace($displayName)) {
        return $displayName
    }

    return $identifier
}

function Copy-OrderedDictionary {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Specialized.OrderedDictionary]$Source
    )

    $copy = [ordered]@{}
    foreach ($key in $Source.Keys) {
        $copy[$key] = $Source[$key]
    }
    return $copy
}

function Get-UniqueSortedStrings {
    param([Parameter()][object[]]$Values = @())

    $seen = New-Object "System.Collections.Generic.HashSet[string]" (
        [StringComparer]::OrdinalIgnoreCase
    )
    $uniqueValues = New-Object System.Collections.Generic.List[string]
    foreach ($value in $Values) {
        $text = [string]$value
        if (
            -not [string]::IsNullOrWhiteSpace($text) -and
            $seen.Add($text)
        ) {
            $uniqueValues.Add($text)
        }
    }

    return @($uniqueValues | Sort-Object { $_.ToLowerInvariant() })
}

function ConvertTo-DiscoveryWorkspaceRows {
param(
    [object[]]$Workspaces,
    [hashtable]$CapacityById,
    $ScanResults,
    [hashtable]$FabricItemCounts,
    [bool]$FabricItemInventoryAvailable = $true,
    [string]$FabricItemInventoryError = ""
)

$reportRows = New-Object System.Collections.Generic.List[object]

foreach ($workspace in $workspaces) {
    $workspaceId = [string](Get-PropertyValue `
        -InputObject $workspace `
        -Name "id" `
        -DefaultValue "")
    $workspaceCapacityId = [string](Get-PropertyValue `
        -InputObject $workspace `
        -Name "capacityId" `
        -DefaultValue "")

    $capacity = $null
    if (-not [string]::IsNullOrWhiteSpace($workspaceCapacityId)) {
        $capacity = $capacityById[$workspaceCapacityId.ToLowerInvariant()]
    }

    $scanDetail = $scanResults.DetailsById[$workspaceId]
    $scanError = [string]$scanResults.ErrorsById[$workspaceId]
    $scanAvailable = (
        $null -ne $scanDetail -and
        [string]::IsNullOrWhiteSpace($scanError)
    )

    $users = @()
    $reports = @()
    $dashboards = @()
    $datasets = @()
    $dataflows = @()
    if ($scanAvailable) {
        $users = @(
            Get-PropertyValue `
                -InputObject $scanDetail `
                -Name "users" `
                -DefaultValue @() |
                Where-Object {
                    $role = [string](Get-PropertyValue `
                        -InputObject $_ `
                        -Name "groupUserAccessRight" `
                        -DefaultValue "")
                    -not [string]::IsNullOrWhiteSpace($role) -and
                    $role -ne "None"
                }
        )
        $reports = @(Get-PropertyValue `
            -InputObject $scanDetail `
            -Name "reports" `
            -DefaultValue @())
        $dashboards = @(Get-PropertyValue `
            -InputObject $scanDetail `
            -Name "dashboards" `
            -DefaultValue @())
        $datasets = @(Get-PropertyValue `
            -InputObject $scanDetail `
            -Name "datasets" `
            -DefaultValue @())
        $dataflows = @(Get-PropertyValue `
            -InputObject $scanDetail `
            -Name "dataflows" `
            -DefaultValue @())
    }

    $adminLabels = @(Get-UniqueSortedStrings -Values @(
        $users |
            Where-Object {
                (Get-PropertyValue `
                    -InputObject $_ `
                    -Name "groupUserAccessRight" `
                    -DefaultValue "") -eq "Admin"
            } |
            ForEach-Object { Get-PrincipalLabel -Principal $_ }
    ))

    $largeStorageFormatModels = @(
        $datasets |
            Where-Object {
                (Get-PropertyValue `
                    -InputObject $_ `
                    -Name "targetStorageMode" `
                    -DefaultValue "") -eq "PremiumFiles"
            } |
            ForEach-Object {
                [string](Get-PropertyValue `
                    -InputObject $_ `
                    -Name "name" `
                    -DefaultValue "")
            }
    )

    $datasetsWithSize = @()
    if ($scanAvailable) {
        $datasetsWithSize = @(
            $datasets | Where-Object {
                $null -ne $_.PSObject.Properties["sizeInBytes"] -and
                $null -ne $_.sizeInBytes
            }
        )
    }
    $sizeDataAvailable = if ($scanAvailable) {
        (
            $datasets.Count -eq 0 -or
            $datasetsWithSize.Count -eq $datasets.Count
        )
    }
    else {
        $false
    }
    $largeModelCount = if ($sizeDataAvailable) {
        @(
            $datasetsWithSize |
                Where-Object { [long]$_.sizeInBytes -ge $LargeModelThresholdBytes }
        ).Count
    }
    else {
        $null
    }

    $fabricItemsCount = if ($fabricItemInventoryAvailable) {
        [int]$fabricItemCounts[$workspaceId]
    }
    else {
        $null
    }

    $discoveryErrors = @(
        @(
            $scanError,
            $fabricItemInventoryError
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
    $workspaceType = [string](Get-PropertyValue `
        -InputObject $workspace `
        -Name "type" `
        -DefaultValue "")
    $supportsDirectWorkspaceRoles = $workspaceType -in @(
        "Workspace",
        "Group"
    )
    $isOrphaned = if (-not $scanAvailable -or -not $supportsDirectWorkspaceRoles) {
        $null
    }
    else {
        $adminLabels.Count -eq 0
    }

    $baseRow = [ordered]@{
        WorkspaceName                 = [string](Get-PropertyValue -InputObject $workspace -Name "name" -DefaultValue "")
        WorkspaceId                   = $workspaceId
        WorkspaceType                 = $workspaceType
        WorkspaceHostingMode          = Get-WorkspaceHostingMode -Workspace $workspace -Capacity $capacity
        CapacityId                    = $workspaceCapacityId
        CapacityName                  = [string](Get-PropertyValue -InputObject $capacity -Name "displayName" -DefaultValue "")
        CapacitySku                   = [string](Get-PropertyValue -InputObject $capacity -Name "sku" -DefaultValue "")
        WorkspaceAdmins               = $adminLabels -join "; "
        PrincipalDisplayName          = ""
        PrincipalEmailOrUpn           = ""
        PrincipalIdentifier           = ""
        PrincipalObjectId             = ""
        PrincipalType                 = ""
        DirectoryUserType             = ""
        WorkspaceRole                 = ""
        State                         = [string](Get-PropertyValue -InputObject $workspace -Name "state" -DefaultValue "")
        IsOrphaned                    = $isOrphaned
        IsReadOnly                    = Get-PropertyValue -InputObject $workspace -Name "isReadOnly"
        IsOnDedicatedCapacity         = Get-PropertyValue -InputObject $workspace -Name "isOnDedicatedCapacity" -DefaultValue $false
        DefaultDatasetStorageFormat   = [string](Get-PropertyValue -InputObject $workspace -Name "defaultDatasetStorageFormat" -DefaultValue "")
        HasWorkspaceLevelSettings     = Get-PropertyValue -InputObject $workspace -Name "hasWorkspaceLevelSettings"
        DashboardsCount               = if ($scanAvailable) { $dashboards.Count } else { $null }
        ReportsCount                  = if ($scanAvailable) { $reports.Count } else { $null }
        DatasetsCount                 = if ($scanAvailable) { $datasets.Count } else { $null }
        LargeModelCount               = $largeModelCount
        LargeModelSizeDataAvailable   = $sizeDataAvailable
        LargeStorageFormatCount       = if ($scanAvailable) { $largeStorageFormatModels.Count } else { $null }
        LargeStorageFormatModel       = if ($scanAvailable) { $largeStorageFormatModels -join "|" } else { "" }
        DataflowsCount                = if ($scanAvailable) { $dataflows.Count } else { $null }
        FabricItemsCount              = $fabricItemsCount
        DiscoveryStatus               = if ($discoveryErrors.Count -eq 0) { "Complete" } else { "Partial" }
        DiscoveryErrors               = $discoveryErrors -join "; "
    }

    if ($users.Count -eq 0) {
        $reportRows.Add([PSCustomObject]$baseRow)
        continue
    }

    foreach ($user in $users) {
        $row = Copy-OrderedDictionary -Source $baseRow
        $profile = Get-PropertyValue -InputObject $user -Name "profile"
        $identityObject = if ($null -ne $profile) { $profile } else { $user }
        $emailAddress = if ($null -ne $profile) {
            ""
        }
        else {
            [string](Get-PropertyValue `
                -InputObject $user `
                -Name "emailAddress" `
                -DefaultValue "")
        }
        $identifierProperty = if ($null -ne $profile) { "id" } else { "identifier" }
        $identifier = [string](Get-PropertyValue `
            -InputObject $identityObject `
            -Name $identifierProperty `
            -DefaultValue "")

        $row["PrincipalDisplayName"] = [string](Get-PropertyValue `
            -InputObject $identityObject `
            -Name "displayName" `
            -DefaultValue "")
        $row["PrincipalEmailOrUpn"] = if (-not [string]::IsNullOrWhiteSpace($emailAddress)) {
            $emailAddress
        }
        else {
            $identifier
        }
        $row["PrincipalIdentifier"] = $identifier
        $row["PrincipalObjectId"] = [string](Get-PropertyValue `
            -InputObject $identityObject `
            -Name $(if ($null -ne $profile) { "id" } else { "graphId" }) `
            -DefaultValue "")
        $row["PrincipalType"] = Get-PrincipalTypeLabel -Principal $user
        $row["DirectoryUserType"] = [string](Get-PropertyValue `
            -InputObject $user `
            -Name "userType" `
            -DefaultValue "")
        $row["WorkspaceRole"] = [string](Get-PropertyValue `
            -InputObject $user `
            -Name "groupUserAccessRight" `
            -DefaultValue "")
        $reportRows.Add([PSCustomObject]$row)
    }
}

@(
    $reportRows |
        Sort-Object `
            @{ Expression = { $_.CapacityName.ToLowerInvariant() } },
            @{ Expression = { $_.WorkspaceName.ToLowerInvariant() } },
            @{ Expression = { $_.WorkspaceRole.ToLowerInvariant() } },
            @{ Expression = { $_.PrincipalDisplayName.ToLowerInvariant() } }
)

}
