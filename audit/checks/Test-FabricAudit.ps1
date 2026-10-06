#requires -Version 5.1
<#
.SYNOPSIS
Runs synthetic, network-disabled checks in either Windows PowerShell or pwsh.
.DESCRIPTION
All authentication and HTTP boundaries are replaced before invoking the package.
No tenant identity, credentials, real resource or network connection is used.
#>
[CmdletBinding()]
param()
Set-StrictMode -Version 2.0
$ErrorActionPreference = "Stop"
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))
$temporary = Join-Path ([IO.Path]::GetTempPath()) ("FabricAuditOffline-" + [Guid]::NewGuid().ToString())
$null = New-Item -ItemType Directory -Path $temporary
$module = Import-Module (Join-Path $root "audit\FabricAudit.psm1") -Force -PassThru
$script:assertions = 0
$expectedCollectors = @("Workspaces","Items","ReportAccess","AppAccess","AppContent","ActivityEvents")
$expectedTabs = @("Workspace Access","Summary","Workspaces","Items","Report Access","Apps",
    "App Access","App Content","Activity Events","Collection Notes","Column Guide")
$expectedColumns = @("WorkspaceName","WorkspaceId","WorkspaceType","WorkspaceHostingMode",
    "CapacityId","CapacityName","CapacitySku","WorkspaceAdmins","PrincipalDisplayName","PrincipalEmailOrUpn",
    "PrincipalIdentifier","PrincipalObjectId","PrincipalType","DirectoryUserType","WorkspaceRole",
    "State","IsOrphaned","IsReadOnly","IsOnDedicatedCapacity","DefaultDatasetStorageFormat",
    "HasWorkspaceLevelSettings","DashboardsCount","ReportsCount","DatasetsCount","LargeModelCount",
    "LargeModelSizeDataAvailable","LargeStorageFormatCount","LargeStorageFormatModel",
    "DataflowsCount","FabricItemsCount","DiscoveryStatus","DiscoveryErrors")
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
    $script:assertions++
}
function Assert-Throws {
    param([scriptblock]$Action, [string]$Message)
    $caught = $false
    try { & $Action | Out-Null } catch { $caught = $true }
    Assert-True $caught $Message
}
function Read-Run {
    param($Result)
    Get-Content -Raw -LiteralPath $Result.ManifestPath | ConvertFrom-Json
}
function Invoke-OfflineLauncher {
    param([hashtable]$Parameters)
    # Keep the already-mocked module when the real launcher imports it.
    function Import-Module {
        param($Name, [switch]$Force, $ErrorAction)
        if ($Name -ne (Join-Path $root "audit\FabricAudit.psm1")) {
            throw "Unexpected launcher import in offline test."
        }
    }
    $global:LASTEXITCODE = 0
    $console = @(& (Join-Path $root "Invoke-FabricDiscovery.ps1") @Parameters 2>&1 3>&1 6>&1)
    [PSCustomObject]@{
        ExitCode = $global:LASTEXITCODE
        Console = ($console | Out-String)
        Result = (& $module {
            if ($script:Audit) {
                [PSCustomObject]@{ Status = $script:Audit.Workbook.Status
                    ManifestPath = Join-Path $script:Audit.Directory "RunManifest.json"
                    WorkbookPath = $script:Audit.Workbook.Path }
            }
        })
    }
}
function Assert-Workbook {
    param($Result)
    $m = Read-Run $Result
    Assert-True ($m.Workbook.Status -eq "Complete" -and (Test-Path -LiteralPath $Result.WorkbookPath)) "Launcher creates the workbook automatically"
    Assert-True ((Split-Path (Split-Path $Result.ManifestPath) -Leaf) -eq "Supporting files") "Supporting files do not clutter the workbook folder"
    $book = Open-ExcelPackage -Path $Result.WorkbookPath
    try {
        $names = @($book.Workbook.Worksheets | ForEach-Object Name)
        Assert-True ($names[-1] -eq "Column Guide") "Column Guide is last"
        $notes = 0
        foreach ($sheet in $book.Workbook.Worksheets) {
            Assert-True (-not $sheet.View.ShowGridLines) "Professional presentation hides gridlines"
            Assert-True ($sheet.View.Panes.Count -gt 0) "Every worksheet has frozen panes"
            foreach ($table in $sheet.Tables) {
                Assert-True ($table.ShowFilter -and $table.ShowRowStripes) "Tables keep filters and striped rows"
                for ($c = $table.Address.Start.Column; $c -le $table.Address.End.Column; $c++) {
                    $header = $sheet.Cells[$table.Address.Start.Row,$c]
                    Assert-True ($null -ne $header.Comment -and $header.Comment.Text.Length -gt 10) "Every header has a description"
                    $notes++
                }
            }
            foreach ($cell in $sheet.Cells[$sheet.Dimension.Address]) {
                Assert-True (-not $cell.Formula) "Source values never become executable workbook formulas"
            }
        }
        $guide = $book.Workbook.Worksheets["Column Guide"]
        Assert-True (($guide.Dimension.End.Row - 4) -eq $notes) "Every header note is also in Column Guide"
        if ("WorkspaceAccess.csv" -in $m.Files.Name) {
            Assert-True ($names[0] -eq "Workspace Access") "Original customer report stays first"
            $sheet = $book.Workbook.Worksheets["Workspace Access"]
            $csv = @(Import-Csv -LiteralPath (Join-Path (Split-Path $Result.ManifestPath) "WorkspaceAccess.csv"))
            $columns = @($m.Files | Where-Object Name -eq "WorkspaceAccess.csv")[0].Columns
            Assert-True ($columns.Count -eq 32 -and $sheet.Dimension.End.Column -eq 32) "Original 32-column contract preserved"
            Assert-True (($columns -join "|") -ceq ($expectedColumns -join "|")) "Headers match the frozen original schema"
            for ($c = 1; $c -le $columns.Count; $c++) {
                Assert-True ($sheet.Cells[1,$c].Value -ceq $columns[$c-1]) "Original column names and order"
                for ($r = 0; $r -lt $csv.Count; $r++) {
                    Assert-True ([string]$sheet.Cells[($r+2),$c].Value -ceq [string]$csv[$r].($columns[$c-1])) "Original workspace values preserved"
                }
            }
        }
    } finally { $book.Dispose() }
}

try {
    & $module {
        function script:Invoke-RestMethod { throw "NETWORK IS PROHIBITED in offline checks." }
        function script:Invoke-WebRequest { throw "NETWORK IS PROHIBITED in offline checks." }
        function script:Connect-AzAccount { throw "AUTHENTICATION IS PROHIBITED in offline checks." }
        function script:Get-AzAccessToken { throw "TOKEN ACQUISITION IS PROHIBITED in offline checks." }
        $script:SavedAccountPrerequisites = ${function:Assert-AuditAccountPrerequisites}
        function script:Assert-AuditAccountPrerequisites { }
        $script:C = "11111111-1111-1111-1111-111111111111"
        $script:W = "22222222-2222-2222-2222-222222222222"
        $script:P = "33333333-3333-3333-3333-333333333333"
        $script:R = "44444444-4444-4444-4444-444444444444"
        $script:D = "55555555-5555-5555-5555-555555555555"
        $script:A = "66666666-6666-6666-6666-666666666666"
        $script:O = "77777777-7777-7777-7777-777777777777"
        $script:U = "88888888-8888-8888-8888-888888888888"
        $script:N = "99999999-9999-9999-9999-999999999999"
        $script:TestMode = ""
        $script:Calls = New-Object 'System.Collections.Generic.List[object]'
        $script:Sleeps = New-Object 'System.Collections.Generic.List[double]'
        $script:Refreshes = 0
        $script:Faults = 0
        $script:ExpireOnHeader = $false
        $script:People = @(
            @{ displayName = "Example User"; identifier = "user@example.com"
                emailAddress = "user@example.com"; graphId = $U; principalType = "User"
                userType = "Member"; groupUserAccessRight = "Admin"; reportUserAccessRight = "Read"; appUserAccessRight = "Read" },
            @{ displayName = "Example Group"; identifier = $N; graphId = $N
                principalType = "Group"; groupUserAccessRight = "Viewer"; reportUserAccessRight = "Read"; appUserAccessRight = "Read" }
        )
        $script:WorkspaceData = @(
            @{ id = $W; name = "Example workspace"; capacityId = $C; type = "Workspace"
                state = "Active"; isOnDedicatedCapacity = $true; isReadOnly = $false
                hasWorkspaceLevelSettings = $false; defaultDatasetStorageFormat = "Small" },
            @{ id = $P; name = "Example personal"; type = "PersonalGroup"; state = "Active"
                isOnDedicatedCapacity = $false }
        )
        $script:ItemData = @(
            @{ id = $R; workspaceId = $W; type = "Report"; name = "Example report" },
            @{ id = $D; workspaceId = $W; type = "SemanticModel"; name = "Example model" },
            @{ id = $O; workspaceId = $W; type = "OrgApp"; name = "Example org app" },
            @{ id = $N; workspaceId = $W; type = "Notebook"; name = "Example notebook" }
        )
        function script:Json-Object { param($Value) $Value | ConvertTo-Json -Depth 30 | ConvertFrom-Json }
        $script:AuthCalls = 0
        function script:Connect-AuditAccount { $script:AuthCalls++ }
        $script:SavedTokenHeader = ${function:Get-ApiAuthorizationHeader}
        function script:Get-ApiAuthorizationHeader { param($Audience, [switch]$ForceRefresh) @{ Authorization = "Bearer SYNTHETIC" } }
        $script:OriginalTestHeader = ${function:Get-ApiAuthorizationHeader}
        function script:Get-ApiAuthorizationHeader {
            param($Audience, [switch]$ForceRefresh)
            if ($ForceRefresh) { $script:Refreshes++ }
            if ($script:ExpireOnHeader) { $script:Audit.Deadline = [DateTimeOffset]::UtcNow.AddSeconds(-1) }
            & $script:OriginalTestHeader $Audience
        }
        function script:Start-Sleep {
            param($Seconds, $Milliseconds)
            $script:Sleeps.Add(([double]$Seconds + [double]$Milliseconds / 1000))
        }
        function script:Invoke-AuditHttp {
            param($Parameters)
            if ($Parameters.Headers.Authorization -ne "Bearer SYNTHETIC") { throw "Real auth is prohibited in offline checks." }
            $script:Calls.Add([PSCustomObject]@{ Uri = $Parameters.Uri; Method = $Parameters.Method; Body = $Parameters["Body"] })
            if ($Parameters.MaximumRedirection -ne 0 -or $Parameters.TimeoutSec -gt 90 -or $Parameters.TimeoutSec -lt 1) {
                throw "Missing transport safeguards."
            }
            $uri = [Uri]$Parameters.Uri
            $path = $uri.AbsolutePath
            if ($script:TestMode -match '^Http(\d+)(Once)?$') {
                $statusCode = [int]$Matches[1]
                if ($script:TestMode -notmatch 'Once$' -or $script:Faults -eq 0) {
                    $script:Faults++
                    $error = New-Object System.Exception("SYNTHETIC-SENSITIVE-ERROR")
                    $error | Add-Member -NotePropertyName Response -NotePropertyValue (
                        [PSCustomObject]@{ StatusCode = $statusCode; Headers = @{ "Retry-After" = "7" } })
                    throw $error
                }
            }
            if ($path -eq "/v1.0/myorg/admin/capacities") {
                if ($script:TestMode -eq "CapacityDenied") { throw "Synthetic capacity denial." }
                return Json-Object @{ value = @(@{ id = $C; displayName = "Example F"; sku = "F4" }) }
            }
            if ($path -eq "/v1.0/myorg/admin/groups") {
                if ($script:TestMode -eq "Empty") { return Json-Object @{ value = @() } }
                $skip = [int][regex]::Match($uri.Query, '\$skip=(\d+)').Groups[1].Value
                $top = [int][regex]::Match($uri.Query, '\$top=(\d+)').Groups[1].Value
                return Json-Object @{ value = @($WorkspaceData | Select-Object -Skip $skip -First $top) }
            }
            if ($path -eq "/v1.0/myorg/admin/workspaces/getInfo") {
                $script:LastScan = ($Parameters.Body | ConvertFrom-Json).workspaces
                return Json-Object @{ id = $N }
            }
            if ($path -match '/scanStatus/') {
                if ($script:TestMode -eq "ScanFailure") { return Json-Object @{ status = "Failed"; error = @{ message = "Synthetic scan failure" } } }
                return Json-Object @{ status = "Succeeded" }
            }
            if ($path -match '/scanResult/') {
                $details = @()
                foreach ($id in $script:LastScan) {
                    if ($id -eq $W) {
                        $report = @{ id = $R; name = "Example report"; datasetId = $D }
                        if ($script:TestMode -ne "ReportFallback") { $report.users = $People }
                        if ($script:TestMode -eq "PrincipalKinds") {
                            $report.users = @($People) + @(
                                @{ identifier = "app"; principalType = "App"; reportUserAccessRight = "Read" },
                                @{ identifier = "tenant"; principalType = "None"; reportUserAccessRight = "Read" },
                                @{ identifier = "none"; principalType = "User"; reportUserAccessRight = "None" }
                            )
                        }
                        $details += @{ id = $W; type = "Workspace"; users = $People
                            reports = @($report); dashboards = @(); dataflows = @()
                            datasets = @(@{ id = $D; name = "Example model"; targetStorageMode = "PremiumFiles" }) }
                        if ($script:TestMode -eq "NullWorkspaceUsers") { $details[-1].users = $null }
                    } else { $details += @{ id = $P; type = "PersonalGroup"; users = @(); reports = @(); datasets = @() } }
                }
                return Json-Object @{ workspaces = $details }
            }
            if ($path -eq "/v1/admin/items") {
                if ($script:TestMode -eq "MalformedItems") { return Json-Object @{ wrong = @() } }
                if ($uri.Query -eq "") {
                    $next = "https://api.fabric.microsoft.com/v1/admin/items?continuationToken=second"
                    if ($script:TestMode -eq "HostileContinuation") { $next = "https://graph.microsoft.com/v1.0/groups" }
                    return Json-Object @{ itemEntities = @($ItemData[0..1]); continuationUri = $next }
                }
                if ($script:TestMode -eq "RepeatContinuation") {
                    return Json-Object @{ itemEntities = @(); continuationUri = $Parameters.Uri }
                }
                return Json-Object @{ itemEntities = @($ItemData[2..3]) }
            }
            if ($path -eq "/v1.0/myorg/admin/apps") {
                return Json-Object @{ value = @(@{ id = $A; name = "Example classic app"; workspaceId = $W }) }
            }
            if ($path -match '/admin/(apps|reports)/[^/]+/users$') {
                return Json-Object @{ value = $People }
            }
            if ($path -eq "/v1.0/myorg/apps/$A/reports") {
                if ($script:TestMode -eq "AppContentDenied") { throw "Synthetic published reports denial." }
                return Json-Object @{ value = @(@{ id = $R; name = "Example published report"; datasetId = $D }) }
            }
            if ($path -eq "/v1.0/myorg/apps/$A/dashboards") {
                if ($script:TestMode -eq "AppContentDenied") { throw "Synthetic published dashboards denial." }
                return Json-Object @{ value = @() }
            }
            if ($path -eq "/v1.0/myorg/admin/activityevents") {
                if ($uri.Query -notmatch 'continuation') {
                    $dateText = [regex]::Match([Uri]::UnescapeDataString($uri.Query), '\d{4}-\d{2}-\d{2}').Value
                    $script:ActivityDay = [DateTime]::ParseExact($dateText, "yyyy-MM-dd", [Globalization.CultureInfo]::InvariantCulture)
                }
                $eventId = if ($uri.Query -match 'continuation') { $N } else { $R }
                if ($script:TestMode -eq "ActivityDuplicate") { $eventId = $R }
                $eventId = $eventId.Substring(0,34) + $script:ActivityDay.ToString("dd")
                $event = @{ Id = $eventId; CreationTime = $script:ActivityDay.ToString("yyyy-MM-ddT12:00:00")
                    UserId = "user@example.com"; Operation = "ViewReport"; WorkspaceId = $W
                    CapacityId = $C; ReportId = $R; IsSuccess = $true }
                if ($script:TestMode -eq "ActivityNoCapacity") { $event.Remove("CapacityId") }
                if ($script:TestMode -eq "ActivityOutOfDay") { $event.CreationTime = $script:ActivityDay.AddDays(1).ToString("s") }
                $response = @{ activityEventEntities = @($event) }
                if ($script:TestMode -eq "ActivityEmpty") { return Json-Object @{ activityEventEntities = @() } }
                if ($uri.Query -notmatch 'continuation') {
                    $response.continuationUri = "https://api.powerbi.com/v1.0/myorg/admin/activityevents?continuationToken=second"
                }
                if ($script:TestMode -in @("ActivityRegional","ActivityTokenOnly","ActivityRawToken","ActivityRepeatedToken") -and
                    ($uri.Query -notmatch 'continuation' -or $script:TestMode -eq "ActivityRepeatedToken")) {
                    $response.continuationUri = "https://wabi-test-redirect.analysis.windows.net/v1.0/myorg/admin/activityevents?continuationToken=unused"
                    $response.continuationToken = "%2BRID%3ASynthetic%3D%23RT%3A1"
                    if ($script:TestMode -eq "ActivityRawToken") { $response.continuationToken = "+RID:Synthetic=#RT:1" }
                    if ($script:TestMode -eq "ActivityTokenOnly") { $response.Remove("continuationUri") }
                }
                if ($script:TestMode -eq "ActivityHostileUri" -and $uri.Query -notmatch 'continuation') {
                    $response.continuationUri = "https://graph.microsoft.com/v1.0/groups"
                }
                return Json-Object $response
            }
            throw "Unrecognised synthetic route. Network is never used: $path"
        }
    }

    $tokenCheck = & $module {
        $script:SyntheticTokenRequests = 0
        function Get-AzAccessToken {
            param($ResourceUrl, $ErrorAction)
            $script:SyntheticTokenRequests++
            [PSCustomObject]@{
                Token = ConvertTo-SecureString "offline-token" -AsPlainText -Force
                ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
        }
        try {
            $first = & $script:SavedTokenHeader PowerBI
            $cached = & $script:SavedTokenHeader PowerBI
            $forced = & $script:SavedTokenHeader PowerBI -ForceRefresh
            [PSCustomObject]@{
                HeaderValid = ($first.Authorization -ceq ("Bearer " + "offline-token"))
                CacheValid = ($cached.Authorization -ceq $first.Authorization)
                RefreshValid = ($forced.Authorization -ceq $first.Authorization -and $script:SyntheticTokenRequests -eq 2)
                PlainTokenValid = ((ConvertFrom-TokenValue "offline-token") -ceq "offline-token")
            }
        } finally { $script:TokenCache = @{} }
    }
    Assert-True $tokenCheck.HeaderValid "SecureString tokens become correct authorization headers"
    Assert-True $tokenCheck.CacheValid "Cached headers retain the original token"
    Assert-True $tokenCheck.RefreshValid "ForceRefresh reacquires a token"
    Assert-True $tokenCheck.PlainTokenValid "Older plain-string token representations remain supported"

    $arguments = @{
        TenantId = "example.onmicrosoft.com"; OutputDirectory = $temporary
    }
    $yesterday = [DateTime]::UtcNow.Date.AddDays(-1).ToString("yyyy-MM-dd")
    $launch = Invoke-OfflineLauncher $arguments
    Assert-True ($launch.ExitCode -eq 2) "Actual default launcher exits 2 for known partial scope"
    Assert-True ($launch.Console -match 'Workbook:' -and $launch.Console -match 'Collection Notes') "Launcher prints result path and partial explanation"
    $result = $launch.Result
    $manifest = Read-Run $result
    Assert-True (($manifest.RequestedCollectors -join "|") -ceq ($expectedCollectors -join "|")) "Actual launcher selects all six collectors by default"
    Assert-True ($manifest.HistoryStartUtc -eq $yesterday -and $manifest.HistoryEndUtc -eq $yesterday) "Default history is exactly yesterday UTC"
    Assert-True ($manifest.MaxRequests -eq 1000 -and $manifest.MaxRunMinutes -eq 30) "Default run is bounded"
    Assert-True ($manifest.Status -eq "Partial") (
        "API limitations remain Partial; workbook must succeed: " + ($manifest.Issues | ConvertTo-Json -Depth 5 -Compress))
    $dir = Split-Path $result.ManifestPath -Parent
    Assert-True (Test-Path (Join-Path $dir "WorkspaceAccess.csv")) (
        "Workspace collector completed: " + ($manifest.Issues | ConvertTo-Json -Depth 5 -Compress))
    $access = @(Import-Csv (Join-Path $dir "WorkspaceAccess.csv"))
    Assert-True ($access.Count -eq 3) "Two assignments plus a personal workspace placeholder"
    $workspaceRows = @(Import-Csv (Join-Path $dir "Workspaces.csv"))
    Assert-True ($workspaceRows.Count -eq 2) "Workspace-only output is unique"
    Assert-True ($access[0].PSObject.Properties.Name.Count -eq 32) "Released baseline has exactly 32 columns"
    $normal = @($access | Where-Object WorkspaceId -eq "22222222-2222-2222-2222-222222222222")
    Assert-True ($normal[0].LargeModelCount -eq "" -and $normal[0].LargeModelSizeDataAvailable -eq "False") "Missing size stays unknown"
    $personal = @($access | Where-Object WorkspaceType -eq "PersonalGroup")[0]
    Assert-True ($personal.IsOrphaned -eq "") "Personal workspaces are not classified as orphaned"
    $items = @(Import-Csv (Join-Path $dir "Items.csv"))
    Assert-True ($items.Count -eq 4) "Item pages reconcile without duplicating scanner items"
    Assert-True (@($items | Where-Object Sources -eq "FabricAdminItems|WorkspaceScanner").Count -eq 2) "Source intersections retain provenance"
    $report = @(Import-Csv (Join-Path $dir "ReportAccessFindings.csv"))
    Assert-True ($report.Count -eq 1 -and $report[0].PolicyFinding -eq "NeedsReview") "No direct grant is fabricated"
    Assert-True ($report[0].PrincipalEmailOrUpn -eq "user@example.com") "User with existing workspace role is not excluded"
    Assert-True ($report[0].Reason -match "individual is listed" -and
        $report[0].Reason -match "Check whether there is a direct individual grant" -and
        $report[0].Reason -match "not a confirmed policy breach") "Individual finding explains the trigger, review action and uncertainty"
    $allReportAccess = @(Import-Csv (Join-Path $dir "ReportAccess.csv"))
    $groupReportAccess = @($allReportAccess | Where-Object PrincipalType -eq "AD Group")
    Assert-True ($groupReportAccess.Count -eq 1 -and $groupReportAccess[0].PolicyFinding -eq "GroupPrincipal" -and
        $groupReportAccess[0].Reason -match "does not flag group entries") "Group reason explains why it is not an individual finding"
    Assert-True ($report[0].Reason -ceq ($allReportAccess | Where-Object PolicyFinding -eq "NeedsReview").Reason) "Report and findings exports use the same explanation"
    $reasonCases = @(
        @{ Type = "Entire Tenant"; Trigger = "entire tenant"; Action = "Review this broad access"; Limit = "not a confirmed policy breach" },
        @{ Type = "Service Principal"; Trigger = "application identity"; Action = "Review it against the policy"; Limit = "does not assess application permissions" },
        @{ Type = "Service Principal Profile"; Trigger = "service principal profile"; Action = "Review it against the policy"; Limit = "does not assess these permissions" },
        @{ Type = "Unknown"; Trigger = "missing or not covered"; Action = "Identify the principal"; Limit = "no policy conclusion" },
        @{ Type = ""; Trigger = "missing or not covered"; Action = "Identify the principal"; Limit = "no policy conclusion" }
    )
    foreach ($case in $reasonCases) {
        $reason = & $module { param($Type) Get-AuditReportAccessReason $Type } $case.Type
        Assert-True ($reason -match $case.Trigger -and $reason -match $case.Action -and
            $reason -match $case.Limit) "Reason explains the classification and follow-up for '$($case.Type)'"
    }
    $apps = @(Import-Csv (Join-Path $dir "Apps.csv"))
    Assert-True (@($apps | Where-Object AppType -eq "OrgApp").Count -eq 1) "OrgApps inventoried separately"
    Assert-True (@($apps | Where-Object AppType -eq "ClassicApp").Count -eq 1) "Classic apps inventoried"
    $content = @(Import-Csv (Join-Path $dir "AppContent.csv"))
    Assert-True ($content.Count -eq 1 -and $content[0].Coverage -eq "CallerVisibleOnly") "App content not substituted with workspace contents"
    $activities = @(Import-Csv (Join-Path $dir "ActivityEvents.csv"))
    Assert-True ($activities.Count -eq 2) "Activity continuation consumed"
    $calls = & $module { $script:Calls.ToArray() }
    Assert-True (@($calls | Where-Object Uri -match '/getInfo').Count -eq 1) "All collectors share one scan"
    Assert-True (@($calls | Where-Object Uri -match '/v1/admin/items').Count -eq 2) "All collectors share item pages"
    $activityCalls = @($calls | Where-Object Uri -match '/admin/activityevents')
    $activityQuery = [Uri]::UnescapeDataString(([Uri]$activityCalls[0].Uri).Query)
    Assert-True ($activityQuery.Contains("'${yesterday}T00:00:00.000Z'") -and
        $activityQuery.Contains("'${yesterday}T23:59:59.999Z'")) "Default request covers one full UTC calendar day"
    Assert-True (@($calls | Where-Object Uri -match 'getDefinition|graph\.|management\.|refreshes').Count -eq 0) "No forbidden services or mutations"
    Assert-True ((Get-Content -Raw $result.ManifestPath) -notmatch 'SYNTHETIC|Bearer') "Manifest excludes tokens"
    Assert-Workbook $result
    $book = Open-ExcelPackage -Path $result.WorkbookPath
    try {
        Assert-True ((@($book.Workbook.Worksheets | ForEach-Object Name) -join "|") -ceq ($expectedTabs -join "|")) "Default workbook has exactly the intended 11 ordered tabs"
        $workspaceSheet = $book.Workbook.Worksheets["Workspaces"]
        $normalRow = @(5..$workspaceSheet.Dimension.End.Row | Where-Object {
            $workspaceSheet.Cells[$_,2].Value -eq "22222222-2222-2222-2222-222222222222"
        })[0]
        Assert-True ($workspaceSheet.Cells[$normalRow,15].Value -eq 0) "Known inventory zero remains numeric zero"
        Assert-True ($null -eq $workspaceSheet.Cells[$normalRow,18].Value) "Unknown model size stays blank"
        Assert-True ($book.Workbook.Worksheets["Report Access"].Cells["N5"].Value -match "policy") "Workbook includes collector reason wording without a separate edit"
        Assert-True ($book.Workbook.Worksheets["Summary"].Cells["A2"].Value -match $yesterday) "Summary uses actual run dates"
    } finally { $book.Dispose() }

    # Frozen expected source text for the original report's normal Admin row.
    $expectedAdmin = @("Example workspace","22222222-2222-2222-2222-222222222222","Workspace","Fabric capacity",
        "11111111-1111-1111-1111-111111111111","Example F","F4","Example User <user@example.com>",
        "Example User","user@example.com","user@example.com","88888888-8888-8888-8888-888888888888",
        "Individual","Member","Admin","Active","False","False","True","Small","False","0","1","1",
        "","False","1","Example model","0","2","Complete","")
    $admin = @($access | Where-Object WorkspaceRole -eq "Admin")[0]
    for ($i = 0; $i -lt $expectedColumns.Count; $i++) {
        Assert-True ($admin.($expectedColumns[$i]) -ceq $expectedAdmin[$i]) "Original rendering: $($expectedColumns[$i])"
    }
    $expectedPersonal = @("Example personal","33333333-3333-3333-3333-333333333333","PersonalGroup",
        "Shared capacity","","","","","","","","","","","","Active","","","False","","",
        "0","0","0","0","True","0","","0","0","Complete","")
    for ($i = 0; $i -lt $expectedColumns.Count; $i++) {
        Assert-True ($personal.($expectedColumns[$i]) -ceq $expectedPersonal[$i]) "Original personal rendering: $($expectedColumns[$i])"
    }
    $expectedGroup = @($expectedAdmin)
    $expectedGroup[8] = "Example Group"; $expectedGroup[9] = "99999999-9999-9999-9999-999999999999"
    $expectedGroup[10] = $expectedGroup[9]; $expectedGroup[11] = $expectedGroup[9]
    $expectedGroup[12] = "AD Group"; $expectedGroup[13] = ""; $expectedGroup[14] = "Viewer"
    $group = @($access | Where-Object WorkspaceRole -eq "Viewer")[0]
    for ($i = 0; $i -lt $expectedColumns.Count; $i++) {
        Assert-True ($group.($expectedColumns[$i]) -ceq $expectedGroup[$i]) "Original group rendering: $($expectedColumns[$i])"
    }
    $sizeCheck = & $module {
        $detail = Json-Object @{ users = @(); reports = @(); dashboards = @(); dataflows = @()
            datasets = @(@{ name = "Below"; sizeInBytes = (1GB - 1) },
                @{ name = "Boundary"; sizeInBytes = 1GB; targetStorageMode = "PremiumFiles" }) }
        $scan = @{ DetailsById = @{ $W = $detail }; ErrorsById = @{} }
        $workspace = Json-Object $WorkspaceData[0]
        @(ConvertTo-DiscoveryWorkspaceRows @($workspace) @{} $scan @{})[0]
    }
    Assert-True ($sizeCheck.LargeModelCount -eq 1 -and $sizeCheck.LargeModelSizeDataAvailable) "Original measured-size threshold remains exactly 1 GiB"
    Assert-True ($sizeCheck.LargeStorageFormatCount -eq 1 -and $sizeCheck.LargeStorageFormatModel -eq "Boundary") "Large storage format remains distinct from measured size"
    Assert-True $sizeCheck.IsOrphaned "A successfully scanned role-assignable workspace without admins is orphaned"
    foreach ($type in @("PersonalGroup","AdminWorkspace")) {
        $row = & $module {
            param($Type)
            $workspace = Json-Object @{ id = $P; name = "Synthetic non-role workspace"; type = $Type; state = "Active" }
            $scan = @{ DetailsById = @{ $P = (Json-Object @{ users = @() }) }; ErrorsById = @{} }
            @(ConvertTo-DiscoveryWorkspaceRows @($workspace) @{} $scan @{})[0]
        } $type
        Assert-True ($row.WorkspaceType -eq $type -and $null -eq $row.IsOrphaned) "Personal/system workspace retained without orphan inference: $type"
    }

    # Per-collector failure and empty-output contracts.
    foreach ($mode in @("Empty","MalformedItems","HostileContinuation","RepeatContinuation","ScanFailure","ReportFallback","NullWorkspaceUsers","CapacityDenied")) {
        & $module { param($Mode) $script:TestMode = $Mode; $script:Calls.Clear() } $mode
        $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
            -Collect Workspaces,Items,ReportAccess -HistoryDays 1 -WorkspacePageSize 1 -WarningAction SilentlyContinue
        $m = Read-Run $r
        Assert-True (Test-Path (Join-Path (Split-Path $r.ManifestPath) "WorkspaceAccess.csv")) "Baseline export retained for $mode"
        if ($mode -eq "Empty") {
            Assert-Workbook $r
            $empty = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "WorkspaceAccess.csv"))
            Assert-True ($empty.Count -eq 0) "Empty extract still has headers"
        }
        if ($mode -eq "HostileContinuation") {
            $evil = @(& $module { $script:Calls | Where-Object Uri -match 'graph\.' })
            Assert-True ($evil.Count -eq 0) "Cross-host continuation rejected before transport"
        }
        if ($mode -eq "ReportFallback") {
            $fallback = @(& $module { $script:Calls | Where-Object Uri -match '/admin/reports/.*/users' })
            Assert-True ($fallback.Count -eq 1) "Missing scanner users are not silently treated as empty"
        }
        if ($mode -eq "NullWorkspaceUsers") {
            $invalid = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "Workspaces.csv") | Where-Object WorkspaceType -eq "Workspace")
            Assert-True ($invalid[0].IsOrphaned -eq "" -and $invalid[0].DiscoveryStatus -eq "Partial") "Null users cannot prove an orphan"
        }
        if ($mode -eq "CapacityDenied") {
            $invalid = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "Workspaces.csv") | Where-Object WorkspaceType -eq "Workspace")
            Assert-True ($invalid[0].CapacityName -eq "" -and $invalid[0].DiscoveryStatus -eq "Partial") "Unavailable capacity metadata is not a complete empty value"
        }
        if ($mode -in @("MalformedItems","HostileContinuation","RepeatContinuation","ScanFailure")) {
            Assert-True (@($m.Issues).Count -gt 1) "Incomplete source errors retained: $mode"
        }
    }

    & $module { $script:TestMode = "PrincipalKinds" }
    $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect ReportAccess -WarningAction SilentlyContinue
    $principals = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "ReportAccess.csv"))
    Assert-True (@($principals | Where-Object PolicyFinding -eq "NotEvaluated").Count -eq 1) "Application access is separately NotEvaluated"
    Assert-True (@($principals | Where-Object PolicyFinding -eq "NeedsReview").Count -eq 2) "Individual and entire-tenant access remain review candidates"
    Assert-True (@($principals | Where-Object Permission -eq "None").Count -eq 0) "No-access principal records are excluded"

    & $module { $script:TestMode = "AppContentDenied" }
    $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect AppAccess,AppContent -MaxRetries 0 -WarningAction SilentlyContinue
    $appNotes = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "AppContentCoverage.csv"))
    Assert-True (@($appNotes | Where-Object Status -eq "Failed").Count -eq 2) "Both denied classic content reads remain Failed"
    Assert-True (@($appNotes | Where-Object Status -eq "Unsupported").Count -eq 1) "OrgApp content remains explicitly unsupported"
    Assert-True ((Read-Run $r).Status -eq "Partial") "A listed app with denied content cannot produce Complete"
    Assert-True (@(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "AppContent.csv")).Count -eq 0) "Denied content is not backfilled from workspace inventory"

    # Exercise the real request wrapper with synthetic HTTP errors, never sleeping.
    foreach ($case in @(
        @{ Mode = "Http429Once"; Requests = 2; Refreshes = 0; Success = $true },
        @{ Mode = "Http503Once"; Requests = 2; Refreshes = 0; Success = $true },
        @{ Mode = "Http401Once"; Requests = 2; Refreshes = 1; Success = $true },
        @{ Mode = "Http401"; Requests = 2; Refreshes = 1; Success = $false },
        @{ Mode = "Http403"; Requests = 1; Refreshes = 0; Success = $false },
        @{ Mode = "Http503"; Requests = 2; Refreshes = 0; Success = $false }
    )) {
        $observed = & $module {
            param($Mode)
            $script:TestMode = $Mode; $script:Faults = 0; $script:Refreshes = 0
            $script:Sleeps.Clear(); $script:Calls.Clear()
            $script:Audit.RequestCount = 0; $script:Audit.MaxRequests = 10
            $script:Audit.Deadline = [DateTimeOffset]::UtcNow.AddMinutes(1)
            $script:Audit.LastRequests = @{}; $script:MaxRetries = 1
            $message = ""; $success = $true
            try { $null = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/capacities" PowerBI -Operation "Synthetic request" -WarningAction SilentlyContinue }
            catch { $success = $false; $message = $_.Exception.Message }
            [PSCustomObject]@{ Requests = $script:Calls.Count; Success = $success; Message = $message
                Refreshes = $script:Refreshes; Sleeps = $script:Sleeps.ToArray() }
        } $case.Mode
        Assert-True ($observed.Requests -eq $case.Requests -and $observed.Success -eq $case.Success) "Retry policy: $($case.Mode)"
        Assert-True ($observed.Refreshes -eq $case.Refreshes) "One token refresh maximum: $($case.Mode)"
        Assert-True ($observed.Message -notmatch "SENSITIVE") "HTTP error text is not persisted"
        if ($case.Mode -eq "Http429Once") { Assert-True (7 -in $observed.Sleeps) "Retry-After honored" }
    }
    & $module {
        $script:TestMode = ""; $script:Calls.Clear(); $script:Sleeps.Clear()
        $script:Audit.LastRequests = @{}; $script:Audit.RequestCount = 0
        $null = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/apps" PowerBI
        $null = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/apps" PowerBI
    }
    $paced = @(& $module { $script:Sleeps.ToArray() })
    Assert-True ($paced.Count -eq 1 -and $paced[0] -gt 17) "200/hour endpoint pacing"
    & $module { $script:Audit.MaxRequests = $script:Audit.RequestCount }
    Assert-Throws { & $module { Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/capacities" PowerBI } } "Request budget is enforced before transport"
    & $module { $script:Audit.Deadline = [DateTimeOffset]::UtcNow.AddSeconds(1) }
    Assert-Throws { & $module { Wait-AuditSeconds 2 } } "Backoff cannot exceed runtime budget"
    & $module {
        $script:Audit.Deadline = [DateTimeOffset]::UtcNow.AddMinutes(1)
        $script:Audit.MaxRequests = 20; $script:Audit.RequestCount = 0
        $script:ExpireOnHeader = $true; $script:Calls.Clear()
    }
    Assert-Throws {
        & $module { Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/capacities" PowerBI }
    } "A slow token acquisition cannot start a request after the deadline"
    $afterAuthCalls = & $module { $script:ExpireOnHeader = $false; $script:Calls.Count }
    Assert-True ($afterAuthCalls -eq 0) "Expired authentication budget prevents transport"

    foreach ($mode in @("ActivityRegional","ActivityTokenOnly","ActivityRawToken","ActivityRepeatedToken","ActivityHostileUri")) {
        & $module { param($Mode) $script:TestMode = $Mode; $script:Calls.Clear() } $mode
        $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
            -Collect ActivityEvents -HistoryDays 1 -WarningAction SilentlyContinue
        $m = Read-Run $r
        $data = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "ActivityEvents.csv"))
        $calls = @(& $module { $script:Calls.ToArray() })
        Assert-True (@($calls | Where-Object { ([Uri]$_.Uri).Host -ne "api.powerbi.com" }).Count -eq 0) "Activity paging never follows another host: $mode"
        if ($mode -eq "ActivityHostileUri") {
            Assert-True ($data.Count -eq 1 -and $m.Status -eq "Partial") "Reject an untrusted continuation without a token"
        } else {
            Assert-True ($data.Count -eq 2 -and $calls.Count -eq 2) "Token pagination consumed once per page: $mode"
            $queryValue = [Uri]::UnescapeDataString(([Uri]$calls[1].Uri).Query.Split("=")[1])
            Assert-True ($queryValue -ceq "'+RID:Synthetic=#RT:1'") "Continuation token is not double-encoded: $mode"
            $expectedStatus = if ($mode -eq "ActivityRepeatedToken") { "Partial" } else { "Complete" }
            Assert-True ($m.Status -eq $expectedStatus) "Repeated-token and completion status: $mode"
        }
    }
    & $module { $script:TestMode = "ActivityNoCapacity" }
    $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect ActivityEvents -CapacityId "11111111-1111-1111-1111-111111111111" -HistoryDays 1 -WarningAction SilentlyContinue
    $gaps = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "ActivityCoverage.csv"))
    Assert-True ($gaps[0].MissingCapacityRows -eq "2") "Historical capacity is not inferred from current workspace placement"

    foreach ($case in @(
        @{ Mode = "ActivityEmpty"; Rows = 0; Status = "Complete" },
        @{ Mode = "ActivityDuplicate"; Rows = 1; Status = "Partial" },
        @{ Mode = "ActivityOutOfDay"; Rows = 0; Status = "Partial" }
    )) {
        & $module { param($Mode) $script:TestMode = $Mode } $case.Mode
        $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary -Collect ActivityEvents
        $data = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "ActivityEvents.csv"))
        Assert-True ($data.Count -eq $case.Rows -and (Read-Run $r).Status -eq $case.Status) "Activity edge case: $($case.Mode)"
    }
    & $module { $script:TestMode = ""; $script:Calls.Clear() }
    $endDay = [DateTime]::UtcNow.Date.AddDays(-2).ToString("yyyy-MM-dd")
    $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect ActivityEvents -HistoryDays 2 -EndDateUtc $endDay
    $rangeManifest = Read-Run $r
    Assert-True ($rangeManifest.HistoryEndUtc -eq $endDay -and
        $rangeManifest.HistoryStartUtc -eq [DateTime]::UtcNow.Date.AddDays(-3).ToString("yyyy-MM-dd")) "Explicit history range is inclusive"
    Assert-True (@(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "ActivityEvents.csv")).Count -eq 4) "Multiple completed days are collected independently"
    Assert-True (@(& $module { $script:Calls | Where-Object Uri -notmatch '/admin/activityevents' }).Count -eq 0) "Activity-only run does not read inventory"

    & $module { $script:TestMode = ""; $script:Calls.Clear() }
    $r = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect Workspaces -CapacityId "11111111-1111-1111-1111-111111111111" -HistoryDays 1 -WarningAction SilentlyContinue
    $filtered = @(Import-Csv (Join-Path (Split-Path $r.ManifestPath) "Workspaces.csv"))
    Assert-True ($filtered.Count -eq 1) "Explicit capacity filtering excludes shared/personal workspace"
    Assert-True ((Read-Run $r).Status -eq "Complete") "Supported complete snapshot can succeed"
    & $module {
        $script:People[0].displayName = "=1+1"
        $script:People[0].identifier = "12345678901234567890"
    }
    $safeText = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect Workspaces,ReportAccess -HistoryDays 1 -WarningAction SilentlyContinue
    $book = Open-ExcelPackage -Path $safeText.WorkbookPath
    try {
        $sheet = $book.Workbook.Worksheets["Report Access"]
        $row = @(5..$sheet.Dimension.End.Row | Where-Object { $sheet.Cells[$_,9].Value -eq "Individual" })[0]
        Assert-True ($sheet.Cells[$row,5].Value -ceq "=1+1" -and -not $sheet.Cells[$row,5].Formula) "Formula-like names remain literal text"
        Assert-True ($sheet.Cells[$row,7].Value -ceq "12345678901234567890") "Long numeric-looking IDs are not rounded"
        Assert-True ($null -eq $book.Workbook.Worksheets["Activity Events"]) "Unselected collector has no misleading empty tab"
    } finally { $book.Dispose() }
    & $module {
        $script:People[0].displayName = "Example User"; $script:People[0].identifier = "user@example.com"
        $script:TestMode = "Http403"
    }
    $failedRead = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect ActivityEvents -HistoryDays 1 -WarningAction SilentlyContinue
    Assert-Workbook $failedRead
    Assert-True ((Read-Run $failedRead).Status -ne "Complete") "A formatted workbook does not disguise a failed API read"
    & $module {
        $script:TestMode = ""
        $script:SavedWorkbookExport = ${function:Export-AuditWorkbook}
        function script:Export-AuditWorkbook { throw "Synthetic workbook write failure." }
    }
    $failedBook = Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary `
        -Collect ActivityEvents -HistoryDays 1 -WarningAction SilentlyContinue
    $failedManifest = Read-Run $failedBook
    Assert-True ($failedManifest.Status -eq "Failed" -and $failedManifest.Workbook.Status -eq "Failed" -and -not $failedBook.WorkbookPath) "Workbook failure is surfaced, not reported as success"
    Assert-True (@($failedManifest.Files | Where-Object Name -eq "CollectionIssues.csv").Count -eq 1) "Rewriting collection issues does not duplicate manifest file entries"
    Assert-True (Test-Path -LiteralPath (Join-Path (Split-Path $failedBook.ManifestPath) "ActivityEvents.csv")) "Workbook failure preserves collected CSVs"
    & $module {
        ${function:Export-AuditWorkbook} = $script:SavedWorkbookExport
        $script:SavedWorkbookPrerequisites = ${function:Assert-AuditWorkbookPrerequisites}
        function script:Assert-AuditWorkbookPrerequisites { throw "Synthetic missing workbook dependency." }
    }
    $authBefore = & $module { $script:AuthCalls }
    Assert-Throws {
        Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -OutputDirectory $temporary -Collect Workspaces
    } "Missing workbook dependency fails before sign-in"
    Assert-True ((& $module { $script:AuthCalls }) -eq $authBefore) "Dependency failure performs no authentication"
    & $module { ${function:Assert-AuditWorkbookPrerequisites} = $script:SavedWorkbookPrerequisites }

    # Exercise the real prerequisite functions with simulated module import failures.
    foreach ($dependency in @("Az.Accounts","ImportExcel")) {
        & $module {
            param($Dependency)
            if ($Dependency -eq "Az.Accounts") {
                function script:Assert-AuditAccountPrerequisites {
                    function Import-Module { throw "Synthetic missing account module." }
                    & $script:SavedAccountPrerequisites
                }
            } else {
                function script:Assert-AuditWorkbookPrerequisites {
                    function Import-Module { throw "Synthetic missing workbook module." }
                    & $script:SavedWorkbookPrerequisites
                }
            }
        } $dependency
        $authBefore = & $module { $script:AuthCalls }
        $launch = Invoke-OfflineLauncher $arguments
        Assert-True ($launch.ExitCode -eq 2 -and $launch.Console -match [regex]::Escape($dependency)) "Actual launcher reports $dependency prerequisite failure"
        Assert-True ((& $module { $script:AuthCalls }) -eq $authBefore) "Missing $dependency never signs in"
        & $module {
            function script:Assert-AuditAccountPrerequisites { }
            ${function:Assert-AuditWorkbookPrerequisites} = $script:SavedWorkbookPrerequisites
        }
    }

    $blockedOutput = Join-Path $temporary "not-a-directory"
    $null = New-Item -ItemType File -Path $blockedOutput
    $authBefore = & $module { $script:AuthCalls }
    $launch = Invoke-OfflineLauncher @{ TenantId = "example.onmicrosoft.com"; OutputDirectory = $blockedOutput }
    Assert-True ($launch.ExitCode -eq 2) "Invalid output directory is a visible launcher failure"
    Assert-True ((& $module { $script:AuthCalls }) -eq $authBefore) "Output directory failure precedes sign-in"

    # Force real IO failures using directories where output files must be written.
    & $module {
        $script:SavedCsvWriter = ${function:Write-AuditCsv}
        $script:BlockedCsvName = ""
        function script:Write-AuditCsv {
            param($Name, [object[]]$Rows, [string[]]$Columns)
            if ($Name -eq $script:BlockedCsvName) {
                $path = Join-Path $script:Audit.Directory $Name
                if (-not (Test-Path -LiteralPath $path)) { $null = New-Item -ItemType Directory -Path $path }
            }
            & $script:SavedCsvWriter @PSBoundParameters
        }
    }
    foreach ($name in @("ActivityEvents.csv","CollectionIssues.csv")) {
        & $module { param($Name) $script:BlockedCsvName = $Name } $name
        $launch = Invoke-OfflineLauncher @{ TenantId = "example.onmicrosoft.com"; OutputDirectory = $temporary; Collect = @("ActivityEvents") }
        $m = Read-Run $launch.Result
        Assert-True ($launch.ExitCode -eq 2 -and $m.Status -eq "Failed") "Actual CSV write failure is not successful: $name"
        Assert-True (@($m.Issues).Count -gt 0) "CSV failure has durable issue context: $name"
        Assert-True ((& $module { $script:TokenCache.Count }) -eq 0) "Token cache clears after output failure"
    }
    & $module {
        ${function:Write-AuditCsv} = $script:SavedCsvWriter
        function script:Export-AuditWorkbook {
            param($CollectionStatus)
            $null = New-Item -ItemType Directory -Path (Join-Path $script:Audit.RunDirectory "Fabric Platform Discovery.xlsx")
            & $script:SavedWorkbookExport $CollectionStatus
        }
    }
    $launch = Invoke-OfflineLauncher @{ TenantId = "example.onmicrosoft.com"; OutputDirectory = $temporary; Collect = @("ActivityEvents") }
    $m = Read-Run $launch.Result
    Assert-True ($launch.ExitCode -eq 2 -and $m.Workbook.Status -eq "Failed" -and -not $launch.Result.WorkbookPath) "Actual workbook target collision is visible with exit 2 and no success path"
    Assert-True ($m.Workbook.Error -match "refusing to overwrite") "Existing workbook target is never overwritten"
    Assert-True (Test-Path -LiteralPath (Join-Path (Split-Path $launch.Result.ManifestPath) "ActivityEvents.csv")) "Workbook IO failure preserves independent CSVs"
    & $module { ${function:Export-AuditWorkbook} = $script:SavedWorkbookExport }

    $launch = Invoke-OfflineLauncher @{ TenantId = "example.onmicrosoft.com"; OutputDirectory = $temporary; Collect = @("Workspaces") }
    Assert-True ($launch.ExitCode -eq 0 -and (Read-Run $launch.Result).Status -eq "Complete") "Actual selective launcher returns zero on Complete"
    $m = Read-Run $launch.Result
    Assert-True (@($m.Collectors | Where-Object Status -eq "NotRequested").Count -eq 5) "Selective run marks every unselected collector NotRequested"
    Assert-True (($m.Files.Name -join "|") -notmatch "ActivityEvents|Apps|ReportAccess") "A later run does not backfill earlier data"
    Assert-True ((Split-Path (Split-Path $launch.Result.ManifestPath) -Leaf) -eq "Supporting files") "Each run keeps supporting files together"
    Assert-True ((Split-Path (Split-Path $launch.Result.WorkbookPath) -Leaf) -match '^\d{8}T\d{6}Z-[0-9a-f-]{36}$') "Run folder is uniquely timestamp/GUID-named"

    $numericPath = Join-Path $temporary "numeric-check.xlsx"
    $culture = [Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo("de-DE")
        & $module {
            param($Path)
            $script:AuditWorkbook = @{ Numeric = @("ReportsCount"); TableNumber = 0 }
            $p = New-Object OfficeOpenXml.ExcelPackage
            try {
                $sheet = Add-AuditWorkbookSheet $p "Values" "2463A6"
                $values = @("0",$null,"10.5","12345678901234567890","0.123456789012345678",
                    "=1+1","+1+1","-1+1","@SUM(A1:A2)")
                $rows = @($values | ForEach-Object { [PSCustomObject]@{ ReportsCount = $_ } })
                $null = Add-AuditWorkbookTable $sheet 1 @("ReportsCount") $rows[0..4] "Workspaces.csv" "Numeric" -NoGuide
                $text = Add-AuditWorkbookSheet $p "Text" "2463A6"
                $null = Add-AuditWorkbookTable $text 1 @("ReportsCount") $rows[5..8] "Workspaces.csv" "Literal text" -KeepText -NoGuide
                $p.SaveAs((New-Object IO.FileInfo($Path)))
            } finally { $p.Dispose(); $script:AuditWorkbook = $null }
        } $numericPath
    } finally { [Threading.Thread]::CurrentThread.CurrentCulture = $culture }
    $book = Open-ExcelPackage -Path $numericPath
    try {
        $sheet = $book.Workbook.Worksheets["Values"]
        Assert-True ($sheet.Cells["A2"].Value -is [double] -and $sheet.Cells["A2"].Value -eq 0) "Numeric zero survives save/reopen"
        Assert-True ($null -eq $sheet.Cells["A3"].Value) "Null stays blank through save/reopen"
        Assert-True ($sheet.Cells["A4"].Value -eq 10.5) "Numeric conversion is locale-independent"
        Assert-True ($sheet.Cells["A5"].Value -ceq "12345678901234567890") "Large numeric values retain exact text"
        Assert-True ($sheet.Cells["A6"].Value -ceq "0.123456789012345678") "Long fractional precision retains exact text"
        foreach ($cell in $book.Workbook.Worksheets["Text"].Cells["A2:A5"]) {
            Assert-True ($cell.Value -is [string] -and -not $cell.Formula) "Formula-like text is literal after save/reopen"
        }
        Assert-Throws { & $module { param($Cell) Set-AuditCellText $Cell ("x" * 32768) } $sheet.Cells["B1"] } "Oversized text is not silently truncated"
        Assert-Throws { & $module { param($Cell) Set-AuditCellText $Cell ([string][char]1) } $sheet.Cells["B2"] } "Unsupported control characters fail visibly"
    } finally { $book.Dispose() }
    foreach ($bad in @("NaN","Infinity","not-a-number")) {
        Assert-Throws {
            & $module {
                param($Value)
                $script:AuditWorkbook = @{ Numeric = @("ReportsCount"); TableNumber = 0 }
                $p = New-Object OfficeOpenXml.ExcelPackage
                try {
                    $sheet = Add-AuditWorkbookSheet $p "Invalid" "2463A6"
                    $null = Add-AuditWorkbookTable $sheet 1 @("ReportsCount") @([PSCustomObject]@{ ReportsCount = $Value }) "Workspaces.csv" "Invalid numeric" -NoGuide
                } finally { $p.Dispose(); $script:AuditWorkbook = $null }
            } $bad
        } "Non-finite/malformed numeric fields fail visibly: $bad"
    }

    $authBefore = & $module { $script:AuthCalls }
    foreach ($parameter in @("MetricsDatasetId","MetricCapacityId")) {
        $invalid = @{ TenantId = "example.onmicrosoft.com"; OutputDirectory = $temporary }
        $invalid[$parameter] = "55555555-5555-5555-5555-555555555555"
        Assert-Throws { & (Join-Path $root "Invoke-FabricDiscovery.ps1") @invalid } "Removed launcher parameter is rejected: $parameter"
        Assert-Throws { Invoke-FabricAuditRun @invalid } "Removed module parameter is rejected: $parameter"
    }
    Assert-Throws { & (Join-Path $root "Invoke-FabricDiscovery.ps1") -TenantId "example.onmicrosoft.com" -Collect Usage } "Removed launcher selector is rejected"
    Assert-Throws { Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -Collect Usage } "Removed module selector is rejected"
    Assert-True ((& $module { $script:AuthCalls }) -eq $authBefore) "Removed arguments cannot authenticate"
    foreach ($file in @("Usage.ps1","Invoke-FabricAudit.ps1")) {
        Assert-True (-not (Test-Path -LiteralPath (Join-Path $root $file)) -and
            -not (Test-Path -LiteralPath (Join-Path $root "audit\$file"))) "No removed runtime file: $file"
    }
    $runtimeFiles = @(Get-Item -LiteralPath (Join-Path $root "Invoke-FabricDiscovery.ps1")) +
        @(Get-ChildItem -LiteralPath (Join-Path $root "audit") -File | Where-Object Extension -in @(".ps1",".psm1",".json"))
    foreach ($file in $runtimeFiles) {
        $source = Get-Content -Raw -LiteralPath $file.FullName
        Assert-True ($source -notmatch '\bCU\b|MetricsDatasetId|MetricCapacityId|executeQueries|Dax|Export-AuditUsage|UsageCoverage|CapacityTimepoints|ItemOperationDailyUsage') "No removed runtime paths or schema in $($file.Name)"
    }
    $defaultFiles = @("WorkspaceAccess.csv","Workspaces.csv","Items.csv","ItemSourceReconciliation.csv",
        "ReportAccess.csv","ReportAccessFindings.csv","ReportAccessCoverage.csv","Apps.csv","AppPrincipals.csv",
        "AppAudienceCoverage.csv","AppContent.csv","AppContentCoverage.csv","ActivityEvents.csv","ActivityCoverage.csv","CollectionIssues.csv")
    Assert-True ((@($manifest.Files.Name | Sort-Object) -join "|") -ceq (@($defaultFiles | Sort-Object) -join "|")) "Default files contain only the approved 15 CSV exports"
    Assert-True ((Get-Content -Raw -LiteralPath $result.ManifestPath) -notmatch 'MetricsDatasetId|MetricCapacityId|UsageCoverage') "Manifest has no removed source state"

    $launcherCommand = Get-Command (Join-Path $root "Invoke-FabricDiscovery.ps1")
    $moduleCommand = Get-Command Invoke-FabricAuditRun
    $publicParameters = @("TenantId","OutputDirectory","Collect","CapacityId","HistoryDays","EndDateUtc",
        "WorkspacePageSize","ScanBatchSize","ScanTimeoutSeconds","MaxRetries","MaxRequests","MaxRunMinutes")
    $runningDoc = Get-Content -Raw -LiteralPath (Join-Path $root "docs\RUNNING.md")
    foreach ($name in $publicParameters) {
        Assert-True ($launcherCommand.Parameters.ContainsKey($name) -and $moduleCommand.Parameters.ContainsKey($name)) "Both entry surfaces support $name"
        Assert-True ($runningDoc.Contains('| `' + $name + '` |')) "Parameter documented: $name"
    }
    foreach ($command in @($launcherCommand,$moduleCommand)) {
        $selectorValues = @($command.Parameters["Collect"].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] })[0].ValidValues
        Assert-True (($selectorValues -join "|") -ceq ($expectedCollectors -join "|")) "CLI ValidateSet matches the documented six selectors"
    }
    $guideDoc = Get-Content -Raw -LiteralPath (Join-Path $root "docs\WORKBOOK-GUIDE.md")
    foreach ($tab in $expectedTabs) { Assert-True ($guideDoc.Contains("| $tab |")) "Default tab documented: $tab" }
    Assert-True ($runningDoc -match '\| `HistoryDays` \| `1` \|' -and $runningDoc -match '\| `EndDateUtc` \| Yesterday UTC \|') "Documentation matches default activity window"
    Assert-True ($guideDoc -match 'Additional 3: historical CU consumption.*DEFERRED' -and $guideDoc -match 'not.*deliver CU') "Requirement 3 is explicitly deferred, not replaced by activity"

    foreach ($uri in @(
        "https://graph.microsoft.com/v1.0/groups",
        "https://management.azure.com/subscriptions",
        "https://api.powerbi.com.evil.example/v1.0/myorg/admin/groups",
        "https://api.powerbi.com:444/v1.0/myorg/admin/groups",
        "https://api.powerbi.com/v1.0/myorg/datasets/55555555-5555-5555-5555-555555555555/executeQueries",
        "https://api.powerbi.com/v1.0/myorg/groups/22222222-2222-2222-2222-222222222222/users",
        "https://api.powerbi.com/v1.0/myorg/admin/tenantSettings",
        "https://api.powerbi.com/v1.0/myorg/admin/groups#redirect",
        "https://user@api.powerbi.com/v1.0/myorg/admin/groups",
        "http://api.powerbi.com/v1.0/myorg/admin/groups",
        "https://api.powerbi.com/v1.0/myorg/datasets/55555555-5555-5555-5555-555555555555/refreshes",
        "https://api.fabric.microsoft.com/v1/workspaces/22222222-2222-2222-2222-222222222222/orgApps/77777777-7777-7777-7777-777777777777/getDefinition"
    )) {
        foreach ($method in @("Get","Post")) {
            Assert-Throws { & $module { param($Uri,$Method) Assert-AuditRequest $Uri PowerBI $Method $null } $uri $method } "Block unapproved $method route: $uri"
        }
    }
    $callsBefore = & $module { $script:Calls.Count }
    Assert-Throws {
        & $module { Invoke-DiscoveryApiRequest "$PowerBIBaseUri/datasets/$D/executeQueries" PowerBI -Method Post -Body @{} }
    } "Model query is blocked through the real transport wrapper"
    Assert-True ((& $module { $script:Calls.Count }) -eq $callsBefore) "Blocked model query cannot reach HTTP"
    & $module {
        Assert-AuditRequest "$FabricBaseUri/admin/items" Fabric Get $null
        Assert-AuditRequest "$PowerBIBaseUri/admin/workspaces/getInfo?getArtifactUsers=true" PowerBI Post @{ workspaces = @($W) }
    }
    Assert-True $true "Read-only Fabric item and scanner routes remain allowed"
    Assert-Throws { & $module { Assert-AuditRequest "$PowerBIBaseUri/admin/groups" PowerBI Get @{} } } "GET request bodies are rejected"
    Assert-Throws {
        Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -Collect Workspaces -OutputDirectory $temporary `
            -EndDateUtc ([DateTime]::UtcNow.ToString("yyyy-MM-dd"))
    } "Reject unfinished/future day before sign-in"
    Assert-Throws {
        Invoke-FabricAuditRun -TenantId "example.onmicrosoft.com" -Collect ActivityEvents -OutputDirectory $temporary `
            -HistoryDays 28 -EndDateUtc ([DateTime]::UtcNow.Date.AddDays(-2).ToString("yyyy-MM-dd"))
    } "Reject an activity range starting beyond retention before sign-in"
    & $module { $script:Audit.Deadline = [DateTimeOffset]::UtcNow.AddSeconds(-1) }
    Assert-Throws { & $module { Assert-AuditTime } } "Runtime guard"

    Write-Host ("PASS: {0} offline assertions on PowerShell {1}. No network/authentication used." -f $script:assertions,$PSVersionTable.PSVersion)
} finally {
    Remove-Module $module -ErrorAction SilentlyContinue
    # This exact GUID-named folder was created above solely for synthetic outputs.
    Remove-Item -LiteralPath $temporary -Recurse -Force
}
