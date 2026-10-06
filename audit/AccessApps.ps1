# 3. Report access is evidence, not a complete description of grant origin.
function Get-AuditReportAccessReason {
    param([string]$PrincipalType)
    switch ($PrincipalType) {
        "Individual" {
            "An individual is listed with report access, while the policy requires group-based access. Check whether there is a direct individual grant. The API does not distinguish direct from inherited access, so this is not a confirmed policy breach."
        }
        "AD Group" {
            "The listed principal is an AD group, not an individual. This check does not flag group entries. Whether the group is approved is outside this check."
        }
        "Entire Tenant" {
            "Access is recorded for the entire tenant, not a specific approved group. Review this broad access against the groups-only policy. The API does not show how it was granted, so this is not a confirmed policy breach."
        }
        "Service Principal" {
            "This is an application identity (service principal), not a person. The groups-only user-access check does not assess application permissions. Review it against the policy for application/service accounts."
        }
        "Service Principal Profile" {
            "This is a service principal profile used by an application, not a person. The groups-only user-access check does not assess these permissions. Review it against the policy for application/service accounts."
        }
        default {
            "The principal type is missing or not covered by this groups-only check. Identify the principal and review its access separately; no policy conclusion is made."
        }
    }
}

function Export-AuditReportAccess {
    $scan = Get-AuditScan
    $reports = @(Get-AuditItemRows | Where-Object { $_.ItemType -in @("Report","PaginatedReport") })
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $coverage = New-Object 'System.Collections.Generic.List[object]'
    foreach ($report in $reports) {
        $source = "WorkspaceScanner"
        try {
            $detail = $scan.DetailsById[$report.WorkspaceId]
            $matches = @(Get-AuditValue $detail "reports" @() | Where-Object { $_.id -eq $report.ItemId })
            $users = $null
            if ($matches.Count -eq 1 -and -not $scan.ErrorsById.ContainsKey($report.WorkspaceId) -and
                $null -ne $matches[0].PSObject.Properties["users"]) {
                $users = @(Get-AuditArray $matches[0] "users")
            } else {
                $source = "GetReportUsersAsAdmin"
                $response = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/reports/$($report.ItemId)/users" `
                    PowerBI -Operation "Read report users"
                $users = @(Get-AuditArray $response "value")
            }
            foreach ($user in $users) {
                $permission = [string](Get-AuditValue $user "reportUserAccessRight" "")
                if (-not $permission) { throw "Report access record omitted its permission." }
                if ($permission -eq "None") { continue }
                $principal = Get-AuditPrincipal $user
                $finding = switch ($principal.PrincipalType) {
                    "Individual" { "NeedsReview" }
                    "AD Group" { "GroupPrincipal" }
                    "Entire Tenant" { "NeedsReview" }
                    default { "NotEvaluated" }
                }
                $row = [ordered]@{
                    WorkspaceId = $report.WorkspaceId; ReportId = $report.ItemId
                    ReportName = $report.ItemName; Permission = $permission
                }
                foreach ($key in $principal.Keys) { $row[$key] = $principal[$key] }
                $row.Source = $source
                $row.GrantOrigin = "NotExposedBySource"
                $row.PolicyFinding = $finding
                $row.Reason = Get-AuditReportAccessReason $principal.PrincipalType
                $rows.Add([PSCustomObject]$row)
            }
            $coverage.Add([PSCustomObject]@{
                WorkspaceId = $report.WorkspaceId; ReportId = $report.ItemId
                AccessCollection = "Complete"; DirectGrantAssessment = "Unsupported"
                Source = $source
            })
        } catch {
            Add-AuditIssue "ReportAccess" $report.ItemId "AccessReadFailed" $_.Exception.Message
            $coverage.Add([PSCustomObject]@{
                WorkspaceId = $report.WorkspaceId; ReportId = $report.ItemId
                AccessCollection = "Failed"; DirectGrantAssessment = "Unavailable"; Source = $source
            })
        }
    }
    $columns = @("WorkspaceId","ReportId","ReportName","Permission","PrincipalDisplayName",
        "PrincipalEmailOrUpn","PrincipalIdentifier","PrincipalObjectId","PrincipalType",
        "DirectoryUserType","Source","GrantOrigin","PolicyFinding","Reason")
    Write-AuditCsv "ReportAccess.csv" $rows.ToArray() $columns
    Write-AuditCsv "ReportAccessFindings.csv" @($rows | Where-Object PolicyFinding -eq "NeedsReview") $columns
    Write-AuditCsv "ReportAccessCoverage.csv" $coverage.ToArray() @(
        "WorkspaceId","ReportId","AccessCollection","DirectGrantAssessment","Source")
    Add-AuditIssue "ReportAccess" "groups-only policy" "UnsupportedGrantOrigin" (
        "Public report-user/scanner schemas do not expose grant origin. Individuals remain NeedsReview, " +
        "including people who also have workspace/group access; no confirmed violation is inferred.")
}

# 4. Distinguish classic apps, OrgApps and audience items. Metadata existence
# does not establish publication, audience recipients or published content.
function Get-AuditApps {
    if ($null -ne $script:Audit.Apps) { return $script:Audit.Apps }
    $catalogue = Get-AuditCatalogue
    $fabric = Get-AuditFabricItems
    $apps = @{}
    foreach ($item in $fabric.Items) {
        if ($item.type -notin @("App","OrgApp","OrgAppAudience")) { continue }
        $id = Get-AuditId $item.id
        $apps[$id] = [PSCustomObject][ordered]@{
            WorkspaceId = [string]$item.workspaceId; AppId = $id
            AppName = [string](Get-AuditValue $item "name" (Get-AuditValue $item "displayName" ""))
            AppType = [string]$item.type; Source = "FabricAdminItems"
            LastUpdate = ""; PublicationStatus = "NotEstablished"
        }
    }
    $skip = 0
    $seen = @{}
    try {
        do {
            $response = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/apps?`$top=100&`$skip=$skip" `
                PowerBI -Operation "List classic apps"
            $page = @(Get-AuditArray $response "value")
            foreach ($app in $page) {
                $id = Get-AuditId (Get-AuditValue $app "id")
                if ($seen.ContainsKey($id)) { throw "Repeated classic app identity during paging." }
                $seen[$id] = $true
                $wid = [string](Get-AuditValue $app "workspaceId" "")
                if (-not $wid -and $apps.ContainsKey($id)) { $wid = $apps[$id].WorkspaceId }
                if ($wid -and -not $catalogue.Workspaces.ContainsKey($wid)) { continue }
                if (-not $wid) {
                    Add-AuditIssue "Apps" $id "MissingWorkspace" "App source omitted workspace ID; association is unresolved."
                    if ($script:Audit.CapacityFilter.Count -gt 0) { continue }
                }
                $apps[$id] = [PSCustomObject][ordered]@{
                    WorkspaceId = $wid; AppId = $id
                    AppName = [string](Get-AuditValue $app "name" "")
                    AppType = "ClassicApp"; Source = "GetAppsAsAdmin"
                    LastUpdate = [string](Get-AuditValue $app "lastUpdate" "")
                    PublicationStatus = "ListedByAdminAppAPI"
                }
            }
            $skip += $page.Count
        } while ($page.Count -eq 100)
    } catch { Add-AuditIssue "Apps" "classic inventory" "ReadFailed" $_.Exception.Message }
    $script:Audit.Apps = @($apps.Values | Sort-Object WorkspaceId,AppType,AppId)
    $script:Audit.Apps
}

function Write-AuditAppsInventory {
    param([object[]]$Apps)
    if (@($script:Audit.Files | Where-Object Name -eq "Apps.csv").Count -eq 0) {
        Write-AuditCsv "Apps.csv" $Apps @("WorkspaceId","AppId","AppName","AppType","Source","LastUpdate","PublicationStatus")
    }
}

function Export-AuditApps {
    $apps = @(Get-AuditApps)
    Write-AuditAppsInventory $apps
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $coverage = New-Object 'System.Collections.Generic.List[object]'
    foreach ($app in $apps) {
        $status = "Unsupported"
        if ($app.AppType -eq "ClassicApp") {
            try {
                $response = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/admin/apps/$($app.AppId)/users" `
                    PowerBI -Operation "Read classic app users"
                foreach ($user in @(Get-AuditArray $response "value")) {
                    $permission = [string](Get-AuditValue $user "appUserAccessRight" "")
                    if (-not $permission) { throw "App access record omitted its permission." }
                    if ($permission -eq "None") { continue }
                    $row = [ordered]@{
                        WorkspaceId = $app.WorkspaceId; AppId = $app.AppId
                        AppType = $app.AppType; Permission = $permission
                    }
                    $principal = Get-AuditPrincipal $user
                    foreach ($key in $principal.Keys) { $row[$key] = $principal[$key] }
                    $row.Source = "GetAppUsersAsAdmin"
                    $rows.Add([PSCustomObject]$row)
                }
                $status = "Complete"
            } catch {
                $status = "Failed"
                Add-AuditIssue "AppAccess" $app.AppId "AppUsersReadFailed" $_.Exception.Message
            }
        }
        $coverage.Add([PSCustomObject]@{
            WorkspaceId = $app.WorkspaceId; AppId = $app.AppId; AppType = $app.AppType
            AppPrincipalCollection = $status; AudienceMembership = "Unsupported"
            Reason = "App users are not audience members. No read-only audience-recipient route established."
        })
    }
    Write-AuditCsv "AppPrincipals.csv" $rows.ToArray() @(
        "WorkspaceId","AppId","AppType","Permission","PrincipalDisplayName","PrincipalEmailOrUpn",
        "PrincipalIdentifier","PrincipalObjectId","PrincipalType","DirectoryUserType","Source")
    Write-AuditCsv "AppAudienceCoverage.csv" $coverage.ToArray() @(
        "WorkspaceId","AppId","AppType","AppPrincipalCollection","AudienceMembership","Reason")
    Add-AuditIssue "AppAccess" "audiences" "UnsupportedAudienceMembership" (
        "Classic app users lack audience fields. OrgAppAudience definitions lack recipients and require item Write; " +
        "these definition calls are excluded. OrgApp metadata is not recipient mapping.")
}

# 5. Classic app content uses delegated app-read APIs, not tenant-admin elevation.
function Export-AuditAppContent {
    $apps = @(Get-AuditApps)
    Write-AuditAppsInventory $apps
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $coverage = New-Object 'System.Collections.Generic.List[object]'
    foreach ($app in $apps) {
        if ($app.AppType -ne "ClassicApp") {
            $coverage.Add([PSCustomObject]@{
                AppId = $app.AppId; AppType = $app.AppType; ContentType = ""
                Status = "Unsupported"; Rows = 0
                Reason = "No verified content route within item-read-only permissions; metadata is not published content."
            })
            continue
        }
        foreach ($kind in @("reports","dashboards")) {
            $status = "CallerVisibleOnly"
            $count = 0
            $reason = "Delegated caller-visible published list; complete audience/content visibility is not established."
            try {
                $response = Invoke-DiscoveryApiRequest "$PowerBIBaseUri/apps/$($app.AppId)/$kind" `
                    PowerBI -Operation "Read published app $kind"
                $seen = @{}
                foreach ($item in @(Get-AuditArray $response "value")) {
                    $id = Get-AuditId (Get-AuditValue $item "id")
                    if ($seen.ContainsKey($id)) { throw "Duplicate published item ID." }
                    $seen[$id] = $true
                    $rows.Add([PSCustomObject][ordered]@{
                        AppWorkspaceId = $app.WorkspaceId; AppId = $app.AppId
                        AppName = $app.AppName; AppType = $app.AppType
                        PublishedItemId = $id
                        ItemName = [string](Get-AuditValue $item "name" (Get-AuditValue $item "displayName" ""))
                        ItemType = $(if ($kind -eq "reports") { "Report" } else { "Dashboard" })
                        OriginalReportId = [string](Get-AuditValue $item "originalReportId" "")
                        DatasetId = [string](Get-AuditValue $item "datasetId" "")
                        Source = "Apps/$kind"; Coverage = "CallerVisibleOnly"
                    })
                    $count++
                }
            } catch {
                $status = "Failed"
                $reason = $_.Exception.Message
                Add-AuditIssue "AppContent" "$($app.AppId)/$kind" "ContentReadFailed" $reason
            }
            $coverage.Add([PSCustomObject]@{
                AppId = $app.AppId; AppType = $app.AppType; ContentType = $kind
                Status = $status; Rows = $count; Reason = $reason
            })
        }
    }
    Write-AuditCsv "AppContent.csv" $rows.ToArray() @(
        "AppWorkspaceId","AppId","AppName","AppType","PublishedItemId","ItemName","ItemType",
        "OriginalReportId","DatasetId","Source","Coverage")
    Write-AuditCsv "AppContentCoverage.csv" $coverage.ToArray() @("AppId","AppType","ContentType","Status","Rows","Reason")
    Add-AuditIssue "AppContent" "published content" "LimitedContentCoverage" (
        "Classic content is caller-visible only. OrgApp definitions require item Write and are not requested. " +
        "Neither complete tenant app content nor per-audience visibility is claimed.")
}
