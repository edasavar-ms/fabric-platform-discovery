# Completed UTC-day activity history, independent of today's inventory and access.
function ConvertTo-AuditUtc {
    param($Value)
    # PowerShell versions differ in whether JSON ISO timestamps become DateTime.
    if ($Value -is [DateTimeOffset]) { return $Value.ToUniversalTime() }
    if ($Value -is [DateTime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) {
            $Value = [DateTime]::SpecifyKind($Value, [DateTimeKind]::Utc)
        }
        return [DateTimeOffset]::new($Value.ToUniversalTime())
    }
    if ($Value -notmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}') {
        throw "Missing or non-ISO source timestamp."
    }
    [DateTimeOffset]::Parse($Value, [Globalization.CultureInfo]::InvariantCulture,
        ([Globalization.DateTimeStyles]::AssumeUniversal -bor [Globalization.DateTimeStyles]::AdjustToUniversal))
}

function Export-AuditActivities {
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $coverage = New-Object 'System.Collections.Generic.List[object]'
    $ids = @{}
    for ($day = $script:Audit.StartDate; $day -le $script:Audit.EndDate; $day = $day.AddDays(1)) {
        $count = 0; $unscoped = 0; $duplicates = 0; $status = "Complete"; $message = ""
        try {
            if ($day -lt [DateTime]::UtcNow.Date.AddDays(-28)) { throw "Date is outside the Activity Events API lookback." }
            $start = [Uri]::EscapeDataString("'" + $day.ToString("yyyy-MM-ddT00:00:00.000Z") + "'")
            $end = [Uri]::EscapeDataString("'" + $day.ToString("yyyy-MM-ddT23:59:59.999Z") + "'")
            $uri = "$PowerBIBaseUri/admin/activityevents?startDateTime=$start&endDateTime=$end"
            $pages = @{}
            do {
                if ($pages.ContainsKey($uri)) { throw "Repeated activity continuation URI." }
                $pages[$uri] = $true
                $response = Invoke-DiscoveryApiRequest $uri PowerBI -Operation "Read historical activity events"
                foreach ($event in @(Get-AuditArray $response "activityEventEntities")) {
                    $id = Get-AuditId (Get-AuditValue $event "Id")
                    $time = ConvertTo-AuditUtc (Get-AuditValue $event "CreationTime" "")
                    if ($time.UtcDateTime.Date -ne $day.Date) { throw "Activity response contains an out-of-day event." }
                    if ($ids.ContainsKey($id)) { $duplicates++; continue }
                    $ids[$id] = $true
                    $capacity = [string](Get-AuditValue $event "CapacityId" "")
                    if ($script:Audit.CapacityFilter.Count -gt 0) {
                        # Current placement cannot safely classify a historical event.
                        if (-not $capacity) { $unscoped++; continue }
                        if ($capacity -notin $script:Audit.CapacityFilter) { continue }
                    }
                    $rows.Add([PSCustomObject][ordered]@{
                        EventId = $id; CreationTimeUtc = $time.ToString("o")
                        Operation = [string](Get-AuditValue $event "Operation" "")
                        UserId = [string](Get-AuditValue $event "UserId" "")
                        WorkspaceId = [string](Get-AuditValue $event "WorkspaceId" "")
                        CapacityId = $capacity
                        ReportId = [string](Get-AuditValue $event "ReportId" "")
                        DatasetId = [string](Get-AuditValue $event "DatasetId" "")
                        ObjectId = [string](Get-AuditValue $event "ObjectId" "")
                        ItemName = [string](Get-AuditValue $event "ItemName" "")
                        ActivityId = [string](Get-AuditValue $event "ActivityId" "")
                        IsSuccess = Get-AuditValue $event "IsSuccess"
                        Source = "GetActivityEvents"
                    })
                    $count++
                }
                $uri = [string](Get-AuditValue $response "continuationUri" "")
                $token = [string](Get-AuditValue $response "continuationToken" "")
                if ($token) {
                    # Use the approved host; the API may supply an encoded token
                    # alongside a regional continuation URI.
                    $decodedToken = [Uri]::UnescapeDataString($token)
                    $uri = "$PowerBIBaseUri/admin/activityevents?continuationToken=" +
                        [Uri]::EscapeDataString("'$decodedToken'")
                }
            } while ($uri)
            if ($unscoped -gt 0 -or $duplicates -gt 0) {
                $status = "Partial"
                $message = "Some events could not be scoped or were duplicated; see counts."
                Add-AuditIssue "ActivityEvents" $day.ToString("yyyy-MM-dd") "CoverageGap" $message
            }
        } catch {
            $status = "Failed"; $message = $_.Exception.Message
            Add-AuditIssue "ActivityEvents" $day.ToString("yyyy-MM-dd") "ActivityReadFailed" $message
        }
        $coverage.Add([PSCustomObject]@{
            DateUtc = $day.ToString("yyyy-MM-dd"); Status = $status; ExportedRows = $count
            MissingCapacityRows = $unscoped; DuplicateRows = $duplicates; Message = $message
        })
    }
    Write-AuditCsv "ActivityEvents.csv" $rows.ToArray() @("EventId","CreationTimeUtc","Operation","UserId",
        "WorkspaceId","CapacityId","ReportId","DatasetId","ObjectId","ItemName","ActivityId","IsSuccess","Source")
    Write-AuditCsv "ActivityCoverage.csv" $coverage.ToArray() @(
        "DateUtc","Status","ExportedRows","MissingCapacityRows","DuplicateRows","Message")
}
