# Internal token, scanner, classification and original report-rendering helpers.
. (Join-Path $PSScriptRoot "DiscoveryCore.ps1")
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"

function Get-AuditValue {
    param($Object, [string]$Name, $Default = $null)
    Get-PropertyValue -InputObject $Object -Name $Name -DefaultValue $Default
}

function Get-AuditArray {
    param($Object, [string]$Name)
    $property = if ($null -ne $Object) { $Object.PSObject.Properties[$Name] }
    if ($null -eq $property -or $null -eq $property.Value -or
        $property.Value -is [string] -or $property.Value -isnot [System.Collections.IEnumerable]) {
        throw "Expected response collection '$Name' was missing or malformed."
    }
    @($property.Value)
}

function Get-AuditId {
    param($Value)
    $id = [Guid]::Empty
    if (-not [Guid]::TryParse([string]$Value, [ref]$id) -or $id -eq [Guid]::Empty) {
        throw "A required resource identifier was missing or invalid."
    }
    $id.ToString()
}

function Add-AuditIssue {
    param([string]$Collector, [string]$Scope, [string]$Code, [string]$Message)
    $script:Audit.Issues.Add([PSCustomObject][ordered]@{
        Collector = $Collector; Scope = $Scope; Code = $Code; Message = $Message
    })
    Write-Warning "$Collector [$Code] $Scope - $Message"
}

function Write-AuditCsv {
    param([string]$Name, [object[]]$Rows, [string[]]$Columns)
    $path = Join-Path $script:Audit.Directory $Name
    # Explicit headers also distinguish a successfully empty export from no file.
    $header = ($Columns | ForEach-Object { '"' + $_.Replace('"', '""') + '"' }) -join ","
    $lines = @($header)
    if ($Rows.Count -gt 0) {
        $lines = @($Rows | Select-Object -Property $Columns | ConvertTo-Csv -NoTypeInformation)
    }
    [IO.File]::WriteAllLines($path, [string[]]$lines, (New-Object Text.UTF8Encoding($true)))
    foreach ($previous in @($script:Audit.Files.ToArray() | Where-Object Name -eq $Name)) {
        $null = $script:Audit.Files.Remove($previous)
    }
    $script:Audit.Files.Add([PSCustomObject]@{
        Name = $Name; Rows = $Rows.Count; Columns = $Columns
    })
}

function Write-AuditManifest {
    param([string]$Status)
    $manifest = [ordered]@{
        SchemaVersion = "1.0"
        Status = $Status
        RunId = $script:Audit.RunId
        Tenant = $script:Audit.Tenant
        StartedUtc = $script:Audit.Started.ToString("o")
        FinishedUtc = [DateTimeOffset]::UtcNow.ToString("o")
        RequestedCollectors = @($script:Audit.Selected)
        CapacityFilter = @($script:Audit.CapacityFilter)
        HistoryStartUtc = $script:Audit.StartDate.ToString("yyyy-MM-dd")
        HistoryEndUtc = $script:Audit.EndDate.ToString("yyyy-MM-dd")
        RequestCount = $script:Audit.RequestCount
        MaxRequests = $script:Audit.MaxRequests
        MaxRunMinutes = $script:Audit.MaxRunMinutes
        Collectors = $script:Audit.Results.ToArray()
        Workbook = $script:Audit.Workbook
        Files = $script:Audit.Files.ToArray()
        Issues = $script:Audit.Issues.ToArray()
        Limitations = @(
            "Inventory/access are current observations, not a historical permission snapshot.",
            "Memberships are not expanded; report-user responses do not prove grant origin.",
            "App-level users are not audience membership; app content is caller-visible.",
            "Activity events are supplementary history, not a permissions snapshot; missing events are not proof of inactivity.",
            "Collector transport excludes Azure resource, Graph, refresh, write-permission and configuration APIs.",
            "Authentication uses Az.Accounts and can read Azure subscription metadata; SkipContextPopulation does not make it Entra-only."
        )
    }
    $path = Join-Path $script:Audit.Directory "RunManifest.json"
    [IO.File]::WriteAllText($path, ($manifest | ConvertTo-Json -Depth 12),
        (New-Object Text.UTF8Encoding($false)))
}

function Assert-AuditTime {
    if ([DateTimeOffset]::UtcNow -ge $script:Audit.Deadline) {
        throw "Audit runtime budget reached. Remaining collection was not attempted."
    }
}

function Wait-AuditSeconds {
    param([double]$Seconds)
    Assert-AuditTime
    if ([DateTimeOffset]::UtcNow.AddSeconds($Seconds) -ge $script:Audit.Deadline) {
        throw "Required wait exceeds the remaining audit runtime budget."
    }
    if ($Seconds -gt 0) { Start-Sleep -Milliseconds ([int][Math]::Ceiling($Seconds * 1000)) }
}

function Assert-AuditRequest {
    param([string]$Uri, [string]$Audience, [string]$Method, $Body)
    $parsed = $null
    if (-not [Uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$parsed) -or
        $parsed.Scheme -ne "https" -or $parsed.Port -ne 443 -or
        $parsed.UserInfo -or $parsed.Fragment) {
        throw "Blocked non-HTTPS or malformed API destination."
    }
    $guid = '[0-9a-fA-F]{8}(?:-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}'
    $p = $parsed.AbsolutePath
    $allowed = $false
    if ($Audience -eq "PowerBI" -and $parsed.Host -eq "api.powerbi.com") {
        if ($Method -eq "Get") {
            $allowed = $p -match "^/v1\.0/myorg/admin/(capacities|groups|apps|activityevents)$" -or
                $p -match "^/v1\.0/myorg/admin/workspaces/(scanStatus|scanResult)/$guid$" -or
                $p -match "^/v1\.0/myorg/admin/(apps|reports)/$guid/users$" -or
                $p -match "^/v1\.0/myorg/apps/$guid/(reports|dashboards)$"
        } elseif ($Method -eq "Post") {
            $allowed = $p -eq "/v1.0/myorg/admin/workspaces/getInfo"
        }
    } elseif ($Audience -eq "Fabric" -and $parsed.Host -eq "api.fabric.microsoft.com") {
        $allowed = $Method -eq "Get" -and $p -eq "/v1/admin/items"
    }
    if (-not $allowed) { throw "Blocked endpoint or method outside the Fabric audit read allowlist." }
    if ($Method -eq "Get" -and $null -ne $Body) { throw "GET requests cannot contain a body." }
}

function Invoke-AuditHttp {
    param([hashtable]$Parameters)
    # No cross-host redirection with a bearer token, and no indefinite HTTP wait.
    Invoke-RestMethod @Parameters
}

# Every collector and scanner call uses this guarded request boundary.
function Invoke-DiscoveryApiRequest {
    param(
        [string]$Uri, [ValidateSet("PowerBI", "Fabric")][string]$Audience,
        [ValidateSet("Get", "Post")][string]$Method = "Get", $Body,
        [string]$Operation
    )
    Assert-AuditRequest $Uri $Audience $Method $Body
    $attempt = 0
    $refreshed = $false
    do {
        Assert-AuditTime
        if ($script:Audit.RequestCount -ge $script:Audit.MaxRequests) {
            throw "Audit request budget reached."
        }
        # Pace rate-limited reads; shared tenant/user traffic can still cause 429.
        $bucket = if ($Uri -match '/admin/apps(?:[/?]|$)') { "Apps" }
            elseif ($Uri -match '/admin/reports/') { "Reports" }
            elseif ($Uri -match '/admin/activityevents') { "Activities" }
            elseif ($Uri -match '/v1/admin/items') { "FabricItems" }
            elseif ($Uri -match '/admin/groups(?:[?]|$)') { "Groups" }
            elseif ($Uri -match '/admin/workspaces/getInfo') { "ScanSubmit" }
        if ($bucket -and $script:Audit.LastRequests.ContainsKey($bucket)) {
            $interval = switch ($bucket) {
                "Groups" { 72.1 }
                "ScanSubmit" { 7.3 }
                default { 18.1 }
            }
            $wait = $interval - ([DateTimeOffset]::UtcNow - $script:Audit.LastRequests[$bucket]).TotalSeconds
            if ($wait -gt 0) { Wait-AuditSeconds $wait }
        }
        $headers = Get-ApiAuthorizationHeader -Audience $Audience
        Assert-AuditTime
        $attempt++
        $script:Audit.RequestCount++
        if ($bucket) { $script:Audit.LastRequests[$bucket] = [DateTimeOffset]::UtcNow }
        $remaining = [Math]::Floor(($script:Audit.Deadline - [DateTimeOffset]::UtcNow).TotalSeconds)
        if ($remaining -lt 1) { throw "Audit runtime budget reached." }
        $parameters = @{
            Uri = $Uri; Method = $Method
            Headers = $headers
            ErrorAction = "Stop"; MaximumRedirection = 0
            TimeoutSec = [int][Math]::Min(90, $remaining)
        }
        if ($null -ne $Body) {
            $parameters.Body = $Body | ConvertTo-Json -Depth 12
            $parameters.ContentType = "application/json"
        }
        try { return Invoke-AuditHttp $parameters }
        catch {
            $status = Get-HttpStatusCode $_
            if ($status -eq 401 -and -not $refreshed) {
                $null = Get-ApiAuthorizationHeader -Audience $Audience -ForceRefresh
                $refreshed = $true
                continue
            }
            if (($null -eq $status -or $status -in @(408,429,500,502,503,504)) -and
                $attempt -le $script:MaxRetries) {
                $delay = Get-RetryAfterSeconds $_ $attempt
                Write-Warning "${Operation}: transient request failure; retrying after $delay second(s)."
                Wait-AuditSeconds $delay
                continue
            }
            # Do not persist raw HTTP errors/response bodies that can contain PII.
            throw "$Operation failed (HTTP $status; attempt $attempt)."
        }
    } while ($true)
}

function Assert-AuditAccountPrerequisites {
    if ($PSVersionTable.PSVersion.Major -lt 6) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }
    try { Import-Module Az.Accounts -ErrorAction Stop }
    catch {
        throw "Az.Accounts could not be loaded. Install it in this PowerShell edition with: Install-Module Az.Accounts -Scope CurrentUser. No sign-in or collection was started."
    }
}

function Connect-AuditAccount {
    # SkipContextPopulation still performs the SDK's initial read-only Azure
    # subscription lookup, separate from the collector allowlist.
    Disable-AzContextAutosave -Scope Process | Out-Null
    Write-Host "Complete the Microsoft sign-in using a Fabric administrator account."
    Connect-AzAccount -Tenant $script:Audit.Tenant -SkipContextPopulation -ErrorAction Stop | Out-Null
}

function Get-AuditCatalogue {
    if ($null -ne $script:Audit.Catalogue) { return $script:Audit.Catalogue }
    $capacities = @{}
    try {
        $response = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/capacities" PowerBI -Operation "List capacities"
        foreach ($capacity in @(Get-AuditArray $response "value")) {
            $capacities[(Get-AuditId (Get-AuditValue $capacity "id"))] = $capacity
        }
    } catch { Add-AuditIssue "Catalogue" "capacities" "ReadFailed" $_.Exception.Message }
    $workspaces = @{}
    $skip = 0
    do {
        $filter = [Uri]::EscapeDataString("state eq 'Active'")
        $uri = "$PowerBIBaseUri/admin/groups?`$filter=$filter&`$top=$WorkspacePageSize&`$skip=$skip"
        $response = Invoke-DiscoveryApiRequest $uri PowerBI -Operation "List active workspaces"
        $page = @(Get-AuditArray $response "value")
        foreach ($workspace in $page) {
            $id = Get-AuditId (Get-AuditValue $workspace "id")
            if ($workspaces.ContainsKey($id)) { throw "Duplicate workspace ID during paging; retry a fresh snapshot." }
            $workspaces[$id] = $workspace
        }
        $skip += $page.Count
    } while ($page.Count -eq $WorkspacePageSize)
    $included = @{}
    foreach ($id in $workspaces.Keys) {
        $capacity = [string](Get-AuditValue $workspaces[$id] "capacityId" "")
        if ($script:Audit.CapacityFilter.Count -eq 0 -or $capacity -in $script:Audit.CapacityFilter) {
            $included[$id] = $workspaces[$id]
        }
    }
    $script:Audit.Catalogue = @{ Workspaces = $included; Capacities = $capacities }
    $script:Audit.Catalogue
}

function Get-AuditScan {
    if ($null -ne $script:Audit.Scan) { return $script:Audit.Scan }
    $catalogue = Get-AuditCatalogue
    $scan = @{ DetailsById = @{}; ErrorsById = @{} }
    if ($catalogue.Workspaces.Count -gt 0) {
        $scan = Get-ScannedWorkspaceDetails -Workspaces @($catalogue.Workspaces.Values)
    }
    foreach ($id in $catalogue.Workspaces.Keys) {
        if ($scan.ErrorsById.ContainsKey($id) -or -not $scan.DetailsById.ContainsKey($id)) { continue }
        $detail = $scan.DetailsById[$id]
        if ((Get-AuditValue $catalogue.Workspaces[$id] "type") -in @("Workspace","Group")) {
            try { $null = @(Get-AuditArray $detail "users") }
            catch { $scan.ErrorsById[$id] = "Workspace users were missing or malformed; access and orphan state are unknown." }
        }
    }
    foreach ($id in $scan.ErrorsById.Keys) {
        Add-AuditIssue "Scan" $id "IncompleteScan" $scan.ErrorsById[$id]
    }
    $script:Audit.Scan = $scan
    $scan
}

function Get-AuditFabricItems {
    if ($null -ne $script:Audit.FabricItems) { return $script:Audit.FabricItems }
    $catalogue = Get-AuditCatalogue
    $result = @{ Items = (New-Object 'System.Collections.Generic.List[object]'); Complete = $true }
    if ($catalogue.Workspaces.Count -eq 0) {
        $script:Audit.FabricItems = $result
        return $result
    }
    $seen = @{}
    $pages = @{}
    $uri = "$FabricBaseUri/admin/items"
    try {
        do {
            if ($pages.ContainsKey($uri)) { throw "Repeated item continuation URI." }
            $pages[$uri] = $true
            $response = Invoke-DiscoveryApiRequest $uri Fabric -Operation "List Fabric items"
            foreach ($item in @(Get-AuditArray $response "itemEntities")) {
                $workspaceId = Get-AuditId (Get-AuditValue $item "workspaceId")
                $id = Get-AuditId (Get-AuditValue $item "id")
                if (-not $catalogue.Workspaces.ContainsKey($workspaceId)) { continue }
                if ([string]::IsNullOrWhiteSpace([string](Get-AuditValue $item "type" ""))) {
                    throw "Fabric item type was missing; inventory and native item counts are incomplete."
                }
                $key = "$workspaceId/$id"
                if ($seen.ContainsKey($key)) { throw "Duplicate Fabric item identity during paging." }
                $seen[$key] = $true
                $result.Items.Add($item)
            }
            $uri = [string](Get-AuditValue $response "continuationUri" "")
        } while ($uri)
    } catch {
        $result.Complete = $false
        Add-AuditIssue "FabricItems" "inventory" "IncompleteInventory" $_.Exception.Message
    }
    $script:Audit.FabricItems = $result
    $result
}

function Get-AuditPrincipal {
    param($Principal)
    $profile = Get-AuditValue $Principal "profile"
    $identity = if ($null -ne $profile) { $profile } else { $Principal }
    $id = [string](Get-AuditValue $identity $(if ($null -ne $profile) { "id" } else { "identifier" }) "")
    $email = if ($null -ne $profile) { "" } else { [string](Get-AuditValue $Principal "emailAddress" "") }
    [ordered]@{
        PrincipalDisplayName = [string](Get-AuditValue $identity "displayName" "")
        PrincipalEmailOrUpn = $(if ($email) { $email } else { $id })
        PrincipalIdentifier = $id
        PrincipalObjectId = [string](Get-AuditValue $identity $(if ($null -ne $profile) { "id" } else { "graphId" }) "")
        PrincipalType = Get-PrincipalTypeLabel $Principal
        DirectoryUserType = [string](Get-AuditValue $Principal "userType" "")
    }
}

. (Join-Path $PSScriptRoot "Inventory.ps1")
. (Join-Path $PSScriptRoot "AccessApps.ps1")
. (Join-Path $PSScriptRoot "Activities.ps1")
. (Join-Path $PSScriptRoot "Workbook.ps1")

function Invoke-FabricAuditRun {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
        [ValidateNotNullOrEmpty()][string]$OutputDirectory = (Join-Path (Get-Location) "Fabric Platform Discovery"),
        [ValidateSet("Workspaces","Items","ReportAccess","AppAccess","AppContent","ActivityEvents")]
        [string[]]$Collect = @("Workspaces","Items","ReportAccess","AppAccess","AppContent","ActivityEvents"),
        [Guid[]]$CapacityId = @(), [ValidateRange(1,28)][int]$HistoryDays = 1,
        [ValidatePattern('^\d{4}-\d{2}-\d{2}$')][string]$EndDateUtc,
        [ValidateRange(1,5000)][int]$WorkspacePageSize = 500,
        [ValidateRange(1,100)][int]$ScanBatchSize = 100,
        [ValidateRange(30,3600)][int]$ScanTimeoutSeconds = 600,
        [ValidateRange(0,10)][int]$MaxRetries = 5,
        [ValidateRange(1,10000)][int]$MaxRequests = 1000,
        [ValidateRange(1,120)][int]$MaxRunMinutes = 30
    )
    if (-not $TenantId.Trim() -or $Collect.Count -eq 0) { throw "Tenant and collectors are required." }
    if ([Guid]::Empty -in $CapacityId) {
        throw "Capacity identifiers cannot be empty GUIDs."
    }
    $end = [DateTime]::UtcNow.Date.AddDays(-1)
    if ($EndDateUtc) {
        $end = [DateTime]::ParseExact($EndDateUtc, "yyyy-MM-dd",
            [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($end -ge [DateTime]::UtcNow.Date) { throw "Use a completed UTC day, not today or a future date." }
    if ("ActivityEvents" -in $Collect -and $end.AddDays(1 - $HistoryDays) -lt [DateTime]::UtcNow.Date.AddDays(-28)) {
        throw "The entire ActivityEvents range must be within the last 28 completed UTC days."
    }
    Assert-AuditAccountPrerequisites
    Assert-AuditWorkbookPrerequisites
    $runId = [Guid]::NewGuid().ToString()
    $runDirectory = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) (
        [DateTime]::UtcNow.ToString("yyyyMMddTHHmmssZ") + "-" + $runId)
    $directory = Join-Path $runDirectory "Supporting files"
    $null = New-Item -ItemType Directory -Path $directory -ErrorAction Stop
    $script:WorkspacePageSize = $WorkspacePageSize
    $script:ScanBatchSize = $ScanBatchSize
    $script:ScanTimeoutSeconds = $ScanTimeoutSeconds
    $script:MaxRetries = $MaxRetries
    $script:TokenCache = @{}
    $script:Audit = @{
        RunId = $runId; Directory = $directory; RunDirectory = $runDirectory; Tenant = $TenantId
        Workbook = [PSCustomObject]@{ Status = "NotStarted"; Path = ""; Sheets = @(); Error = "" }
        CollectionFinished = $null
        Selected = @($Collect | Select-Object -Unique)
        CapacityFilter = @($CapacityId | ForEach-Object { $_.ToString() } | Select-Object -Unique)
        StartDate = $end.AddDays(1 - $HistoryDays); EndDate = $end
        Started = [DateTimeOffset]::UtcNow
        Deadline = [DateTimeOffset]::UtcNow.AddMinutes($MaxRunMinutes)
        MaxRequests = $MaxRequests; MaxRunMinutes = $MaxRunMinutes; RequestCount = 0
        Catalogue = $null; Scan = $null; FabricItems = $null; Apps = $null
        LastRequests = @{}
        Files = (New-Object 'System.Collections.Generic.List[object]')
        Issues = (New-Object 'System.Collections.Generic.List[object]')
        Results = (New-Object 'System.Collections.Generic.List[object]')
    }
    $overall = "Failed"
    Write-AuditManifest "Running"
    try {
        Connect-AuditAccount
        foreach ($collector in @("Workspaces","Items","ReportAccess","AppAccess","AppContent","ActivityEvents")) {
            if ($collector -notin $script:Audit.Selected) {
                $script:Audit.Results.Add([PSCustomObject]@{ Collector = $collector; Status = "NotRequested" })
                continue
            }
            $status = "Complete"
            Write-Host ("Collecting {0}..." -f $collector)
            try {
                Assert-AuditTime
                switch ($collector) {
                    "Workspaces" { Export-AuditWorkspaces }
                    "Items" { Export-AuditItems }
                    "ReportAccess" { Export-AuditReportAccess }
                    "AppAccess" { Export-AuditApps }
                    "AppContent" { Export-AuditAppContent }
                    "ActivityEvents" { Export-AuditActivities }
                }
                $dependencies = switch ($collector) {
                    "Workspaces" { @("Catalogue","Scan","FabricItems") }
                    "Items" { @("Catalogue","Scan","FabricItems") }
                    "ReportAccess" { @("Catalogue","Scan","FabricItems","Items") }
                    "AppAccess" { @("Catalogue","FabricItems","Apps") }
                    "AppContent" { @("Catalogue","FabricItems","Apps") }
                    default { @() }
                }
                if (@($script:Audit.Issues | Where-Object {
                    $_.Collector -eq $collector -or $_.Collector -in $dependencies
                }).Count -gt 0) { $status = "Partial" }
            } catch {
                $status = "Failed"
                Add-AuditIssue $collector "" "CollectorFailed" $_.Exception.Message
            }
            $script:Audit.Results.Add([PSCustomObject]@{ Collector = $collector; Status = $status })
            Write-AuditManifest "Running"
        }
        $requestedResults = @($script:Audit.Results | Where-Object Status -ne "NotRequested")
        $overall = if (@($requestedResults | Where-Object Status -ne "Failed").Count -eq 0) { "Failed" }
        elseif (@($script:Audit.Results | Where-Object {
            $_.Status -notin @("Complete","NotRequested")
        }).Count -gt 0) { "Partial" } else { "Complete" }
    } catch {
        Add-AuditIssue "Run" "" "RunFailed" $_.Exception.Message
        $overall = "Failed"
    } finally {
        $script:Audit.CollectionFinished = [DateTimeOffset]::UtcNow
        try {
            try {
                Write-AuditCsv "CollectionIssues.csv" $script:Audit.Issues.ToArray() @("Collector","Scope","Code","Message")
                $script:Audit.Workbook = Export-AuditWorkbook $overall
            } catch {
                $overall = "Failed"
                $script:Audit.Workbook = [PSCustomObject]@{
                    Status = "Failed"; Path = ""; Sheets = @(); Error = $_.Exception.Message
                }
                Add-AuditIssue "Workbook" "" "WorkbookExportFailed" $_.Exception.Message
                try { Write-AuditCsv "CollectionIssues.csv" $script:Audit.Issues.ToArray() @("Collector","Scope","Code","Message") }
                catch { Add-AuditIssue "Run" "CollectionIssues.csv" "IssueExportFailed" $_.Exception.Message }
            }
            Write-AuditManifest $overall
        } finally { $script:TokenCache = @{} }
    }
    [PSCustomObject]@{
        Status = $overall; ManifestPath = (Join-Path $directory "RunManifest.json")
        WorkbookPath = $script:Audit.Workbook.Path
    }
}

Export-ModuleMember -Function Invoke-FabricAuditRun
