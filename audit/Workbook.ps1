# Format only this run's local exports. No authentication or API calls belong here.
function Assert-AuditWorkbookPrerequisites {
    try { Import-Module ImportExcel -MinimumVersion 7.8.10 -ErrorAction Stop }
    catch {
        throw "Workbook output requires ImportExcel 7.8.10 or later. Install it once with: Install-Module ImportExcel -MinimumVersion 7.8.10 -Scope CurrentUser. No sign-in or collection was started."
    }
    $path = Join-Path $PSScriptRoot "ColumnDefinitions.json"
    $script:AuditColumnDefinitions = Get-Content -Raw -LiteralPath $path -ErrorAction Stop | ConvertFrom-Json
    if ($null -eq $script:AuditColumnDefinitions.defaults -or $null -eq $script:AuditColumnDefinitions.overrides) {
        throw "ColumnDefinitions.json is invalid."
    }
}

function Get-AuditColumnDescription {
    param([string]$Source, [string]$Column)
    $overrides = Get-AuditValue $script:AuditColumnDefinitions.overrides $Source
    $description = Get-AuditValue $overrides $Column
    if (-not $description) { $description = Get-AuditValue $script:AuditColumnDefinitions.defaults $Column }
    if (-not $description) { throw "Column definition missing: $Source/$Column." }
    [string]$description
}

function Set-AuditCellText {
    param($Cell, $Value)
    if ($null -eq $Value) { return }
    $text = [string]$Value
    if ($text.Length -eq 0) { return }
    if ($text.Length -gt 32767 -or $text -match '[\x00-\x08\x0B\x0C\x0E-\x1F]') {
        throw "Source text cannot be represented in an Excel cell without alteration."
    }
    # Value, never Formula: source names/IDs beginning with '=' stay literal text.
    $Cell.Value = $text
    $Cell.Style.Numberformat.Format = "@"
}

function Add-AuditWorkbookSheet {
    param($Package, [string]$Name, [string]$Color)
    $sheet = $Package.Workbook.Worksheets.Add($Name)
    $sheet.View.ShowGridLines = $false
    $sheet.View.ZoomScale = 85
    $sheet.TabColor = [Drawing.ColorTranslator]::FromHtml("#$Color")
    $sheet.DefaultColWidth = 28
    $sheet.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
    $sheet.PrinterSettings.FitToPage = $true
    $sheet.PrinterSettings.FitToWidth = 1
    $sheet.PrinterSettings.FitToHeight = 0
    $sheet
}

function Add-AuditWorkbookBanner {
    param($Sheet, [int]$Row, [string]$Title, [string]$Detail, [int]$Width = 8)
    $Sheet.Cells[$Row,1,$Row,$Width].Merge = $true
    $cell = $Sheet.Cells[$Row,1]
    Set-AuditCellText $cell $Title
    $cell.Style.Font.Name = "Arial"
    $cell.Style.Font.Size = 17
    $cell.Style.Font.Bold = $true
    $cell.Style.Font.Color.SetColor([Drawing.Color]::White)
    $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $cell.Style.Fill.BackgroundColor.SetColor([Drawing.ColorTranslator]::FromHtml("#16324F"))
    $Sheet.Row($Row).Height = 28
    $Sheet.Cells[($Row+1),1,($Row+1),$Width].Merge = $true
    Set-AuditCellText $Sheet.Cells[($Row+1),1] $Detail
    $Sheet.Cells[($Row+1),1].Style.Font.Name = "Arial"
    $Sheet.Cells[($Row+1),1].Style.Font.Size = 10
    $Sheet.Cells[($Row+1),1].Style.WrapText = $true
    $Sheet.Row($Row+1).Height = 36
}

function Add-AuditWorkbookTable {
    param(
        $Sheet, [int]$HeaderRow, [string[]]$Columns, [object[]]$Rows,
        [string]$Source, [string]$Section, [switch]$KeepText, [switch]$NoGuide
    )
    if ($Columns.Count -eq 0 -or $Columns.Count -gt 16384 -or $HeaderRow + [Math]::Max(1,$Rows.Count) -gt 1048576) {
        throw "Export exceeds Excel worksheet limits; CSVs remain available."
    }
    $last = $HeaderRow + [Math]::Max(1,$Rows.Count)
    $range = $Sheet.Cells[$HeaderRow,1,$last,$Columns.Count]
    $range.Style.Font.Name = "Arial"
    $range.Style.Font.Size = 10
    $range.Style.WrapText = $true
    $range.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    for ($column = 1; $column -le $Columns.Count; $column++) {
        $name = $Columns[$column-1]
        $cell = $Sheet.Cells[$HeaderRow,$column]
        Set-AuditCellText $cell $name
        $definition = Get-AuditColumnDescription $Source $name
        $comment = $cell.AddComment($definition, "Fabric Platform Discovery")
        $comment.AutoFit = $true
        if (-not $NoGuide) {
            $script:AuditWorkbook.Guide.Add([PSCustomObject][ordered]@{
                Sheet = $Sheet.Name; Section = $Section; Column = $name; Description = $definition
            })
        }
        $width = if ($name -in @("Reason","Notes","Description","Message","DiscoveryErrors","Representation","What it contains")) { 75 }
            elseif ($name -in @("Files","ReviewContext","WorkspaceAdmins")) { 52 }
            elseif ($name -match 'Id$|Identifier$|Utc$|Name|Email') { 39 }
            else { 28 }
        $Sheet.Column($column).Width = [Math]::Max($Sheet.Column($column).Width, $width)
    }
    $rowNumber = $HeaderRow
    foreach ($row in $Rows) {
        $rowNumber++
        for ($column = 1; $column -le $Columns.Count; $column++) {
            $name = $Columns[$column-1]
            $value = Get-AuditValue $row $name
            $cell = $Sheet.Cells[$rowNumber,$column]
            if (-not $KeepText -and $name -in $script:AuditWorkbook.Numeric -and
                $null -ne $value -and [string]$value -ne "") {
                $number = 0.0
                $text = [Convert]::ToString($value, [Globalization.CultureInfo]::InvariantCulture)
                if (-not [double]::TryParse($text, [Globalization.NumberStyles]::Float,
                    [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -or
                    [double]::IsNaN($number) -or [double]::IsInfinity($number)) {
                    throw "Invalid numeric value in $Source/$name."
                }
                # Keep values beyond Excel's precision as text rather than round them.
                $mantissa = ($text -split '[eE]')[0] -replace '[^0-9]', ''
                $digits = $mantissa.TrimStart('0').TrimEnd('0').Length
                if ($digits -le 15 -and [Math]::Abs($number) -lt 1e15) {
                    $cell.Value = $number
                    $cell.Style.Numberformat.Format = "#,##0.###############"
                } else { Set-AuditCellText $cell $value }
            } else { Set-AuditCellText $cell $value }
            if ([string]$value -in @("NeedsReview","Partial","Unsupported","Failed","NotAttempted","NoRowsReturned")) {
                $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $cell.Style.Fill.BackgroundColor.SetColor([Drawing.ColorTranslator]::FromHtml("#FFF0CB"))
            }
        }
        $Sheet.Row($rowNumber).Height = if ("Reason" -in $Columns -or "Description" -in $Columns) { 66 } else { 42 }
    }
    $script:AuditWorkbook.TableNumber++
    $table = $Sheet.Tables.Add($range, "DiscoveryTable$($script:AuditWorkbook.TableNumber)")
    $table.TableStyle = [OfficeOpenXml.Table.TableStyles]::Medium2
    $table.ShowFilter = $true
    $table.ShowRowStripes = $true
    $Sheet.Row($HeaderRow).Height = 36
    if ($Rows.Count -eq 0) {
        $comment = $Sheet.Cells[($HeaderRow+1),1].AddComment(
            "No data rows were exported. Check Summary and Collection Notes: empty does not imply successful collection.",
            "Fabric Platform Discovery")
        $comment.AutoFit = $true
    }
    $last
}

function Add-AuditWorkbookCsv {
    param($Sheet, [string]$Name, [int]$HeaderRow, [string]$Section, [switch]$KeepText)
    $entries = @($script:Audit.Files | Where-Object Name -eq $Name)
    if ($entries.Count -ne 1) { throw "CSV metadata missing or duplicated: $Name." }
    $entry = $entries[0]
    $path = Join-Path $script:Audit.Directory $Name
    $rows = @(Import-Csv -LiteralPath $path -Encoding UTF8 -ErrorAction Stop)
    if ($rows.Count -ne $entry.Rows) { throw "CSV row count mismatch: $Name." }
    $header = Get-Content -LiteralPath $path -TotalCount 1 -Encoding UTF8
    $expected = ($entry.Columns | ForEach-Object { '"' + $_.Replace('"','""') + '"' }) -join ","
    if ($header -cne $expected) { throw "CSV columns differ from metadata: $Name." }
    $last = Add-AuditWorkbookTable $Sheet $HeaderRow $entry.Columns $rows $Name $Section -KeepText:$KeepText
    $script:AuditWorkbook.Sources.Add([PSCustomObject][ordered]@{
        "Source CSV" = $Name; "Workbook tab" = $Sheet.Name
        "Table header" = "A$HeaderRow"; Records = $rows.Count
        Representation = "Columns and data from this run; IDs and long-precision numbers remain text."
    })
    $last
}

function Export-AuditWorkbook {
    param([string]$CollectionStatus)
    $script:AuditWorkbook = @{
        Guide = (New-Object 'System.Collections.Generic.List[object]')
        Sources = (New-Object 'System.Collections.Generic.List[object]')
        TableNumber = 0
        Numeric = @("DashboardsCount","ReportsCount","DatasetsCount","LargeModelCount",
            "LargeStorageFormatCount","DataflowsCount","FabricItemsCount",
            "Rows","ExportedRows","MissingCapacityRows",
            "DuplicateRows","Records")
    }
    $target = Join-Path $script:Audit.RunDirectory "Fabric Platform Discovery.xlsx"
    $pending = Join-Path $script:Audit.RunDirectory (".workbook-" + [Guid]::NewGuid().ToString() + ".xlsx")
    if (Test-Path -LiteralPath $target) { throw "Workbook already exists; refusing to overwrite it." }
    $package = New-Object OfficeOpenXml.ExcelPackage
    try {
        $names = @($script:Audit.Files | ForEach-Object Name)
        if ("WorkspaceAccess.csv" -in $names) {
            $sheet = Add-AuditWorkbookSheet $package "Workspace Access" "2463A6"
            $null = Add-AuditWorkbookCsv $sheet "WorkspaceAccess.csv" 1 "Original workspace report" -KeepText
            $sheet.View.FreezePanes(2,2)
        }
        $summary = Add-AuditWorkbookSheet $package "Summary" "5B6770"
        $dateRange = $script:Audit.StartDate.ToString("yyyy-MM-dd") + " through " + $script:Audit.EndDate.ToString("yyyy-MM-dd")
        Add-AuditWorkbookBanner $summary 1 "Fabric Platform Discovery" (
            "Collection status: $CollectionStatus. History requested: $dateRange UTC. Inventory/access are current observations, not historical permissions.")
        $settings = @(
            [PSCustomObject]@{ Field = "Tenant"; Value = $script:Audit.Tenant },
            [PSCustomObject]@{ Field = "Run ID"; Value = $script:Audit.RunId },
            [PSCustomObject]@{ Field = "Started UTC"; Value = $script:Audit.Started.ToString("o") },
            [PSCustomObject]@{ Field = "Collection finished UTC"; Value = $script:Audit.CollectionFinished.ToString("o") },
            [PSCustomObject]@{ Field = "Historical range (UTC)"; Value = $dateRange },
            [PSCustomObject]@{ Field = "Collectors requested"; Value = $script:Audit.Selected -join "; " },
            [PSCustomObject]@{ Field = "Capacity filter"; Value = $(if ($script:Audit.CapacityFilter.Count) { $script:Audit.CapacityFilter -join "; " } else { "None (all returned workspaces)" }) },
            [PSCustomObject]@{ Field = "Interpretation"; Value = "Blank is not zero. Partial/unsupported fields remain unavailable. Activity history does not establish current permissions." }
        )
        $last = Add-AuditWorkbookTable $summary 4 @("Field","Value") $settings "Settings" "Run scope"
        Add-AuditWorkbookBanner $summary ($last+3) "Extract status" "NotRequested means that collector was not selected. Partial can include usable data plus unsupported requirements."
        $last = Add-AuditWorkbookTable $summary ($last+6) @("Collector","Status") $script:Audit.Results.ToArray() "Collectors" "Extract status"
        $data = @(
            @("Workspaces","Workspaces.csv","117C83","One row per active workspace, including shared/personal/system workspaces where returned."),
            @("Items","Items.csv","117C83","Workspace/item inventory. Source differences appear in Collection Notes."),
            @("Report Access","ReportAccess.csv","6B5295","Filter PolicyFinding = NeedsReview. Reasons explain what to check; these are not confirmed direct-sharing violations."),
            @("Apps","Apps.csv","6B5295","App and audience metadata types remain distinct. Metadata existence does not prove publication."),
            @("App Access","AppPrincipals.csv","6B5295","Classic app-level principals, not audience membership."),
            @("App Content","AppContent.csv","6B5295","Caller-visible classic published content only. See Collection Notes for denied or unsupported reads."),
            @("Activity Events","ActivityEvents.csv","347A56","Audit activities for the requested UTC range. See coverage for incomplete days. Events do not establish consumption or permission history.")
        )
        foreach ($spec in $data) {
            if ($spec[1] -notin $names) { continue }
            $sheet = Add-AuditWorkbookSheet $package $spec[0] $spec[2]
            Add-AuditWorkbookBanner $sheet 1 $spec[0] $spec[3]
            $null = Add-AuditWorkbookCsv $sheet $spec[1] 4 $spec[0]
            $sheet.View.FreezePanes(5,2)
        }
        if ("ReportAccessFindings.csv" -in $names) {
            # Preserve the findings as a filter, verifying no rows are lost by omitting a duplicate tab.
            if ("ReportAccess.csv" -notin $names) { throw "Report findings have no matching full access export." }
            $all = @(Import-Csv -LiteralPath (Join-Path $script:Audit.Directory "ReportAccess.csv") -Encoding UTF8)
            $findings = @(Import-Csv -LiteralPath (Join-Path $script:Audit.Directory "ReportAccessFindings.csv") -Encoding UTF8)
            $expected = @($all | Where-Object PolicyFinding -eq "NeedsReview")
            if (($expected | ConvertTo-Json -Depth 5 -Compress) -cne ($findings | ConvertTo-Json -Depth 5 -Compress)) {
                throw "Report findings do not match the NeedsReview subset."
            }
            $script:AuditWorkbook.Sources.Add([PSCustomObject][ordered]@{
                "Source CSV" = "ReportAccessFindings.csv"; "Workbook tab" = "Report Access"
                "Table header" = "A4"; Records = $findings.Count
                Representation = "Filter PolicyFinding = NeedsReview. Exact source subset verified; no duplicate tab."
            })
        }
        $notes = Add-AuditWorkbookSheet $package "Collection Notes" "5B6770"
        Add-AuditWorkbookBanner $notes 1 "Collection Notes" "Collection errors, capability limits and source coverage. Empty or failed output is not proof of no access or no activity."
        $sections = @(
            @("CollectionIssues.csv","Collection issues"),
            @("ReportAccessCoverage.csv","Report access coverage"),
            @("AppAudienceCoverage.csv","App principals and audience coverage"),
            @("AppContentCoverage.csv","Published app content coverage"),
            @("ActivityCoverage.csv","Activity coverage"),
            @("ItemSourceReconciliation.csv","Item source differences")
        )
        $row = 5
        foreach ($section in $sections) {
            if ($section[0] -notin $names) { continue }
            Add-AuditWorkbookBanner $notes $row $section[1] ("Source: " + $section[0])
            $row = (Add-AuditWorkbookCsv $notes $section[0] ($row+3) $section[1]) + 4
        }
        $notes.View.FreezePanes(5,2)
        if ($script:AuditWorkbook.Sources.Count -ne $script:Audit.Files.Count) {
            throw "One or more CSV exports have no workbook representation."
        }
        Add-AuditWorkbookBanner $summary ($last+3) "Where the data is" "The supporting CSVs are retained in Supporting files. The original workspace report is first when requested."
        $null = Add-AuditWorkbookTable $summary ($last+6) @("Source CSV","Workbook tab","Table header","Records","Representation") `
            $script:AuditWorkbook.Sources.ToArray() "SourceMap" "Source file map"
        $summary.Column(2).Width = 80
        $summary.View.FreezePanes(5,2)
        $guide = Add-AuditWorkbookSheet $package "Column Guide" "5B6770"
        Add-AuditWorkbookBanner $guide 1 "Column Guide" "Column names are unchanged. Hover over a table header for its note, or read the definitions below." 4
        $guideColumns = @("Sheet","Section","Column","Description")
        foreach ($column in $guideColumns) {
            $script:AuditWorkbook.Guide.Add([PSCustomObject][ordered]@{
                Sheet = "Column Guide"; Section = "Column definitions"; Column = $column
                Description = Get-AuditColumnDescription "ColumnGuide" $column
            })
        }
        $null = Add-AuditWorkbookTable $guide 4 $guideColumns $script:AuditWorkbook.Guide.ToArray() "ColumnGuide" "Column definitions" -NoGuide
        $guide.Column(4).Width = 110
        $guide.View.FreezePanes(5,4)
        $package.Workbook.Properties.Title = "Fabric Platform Discovery"
        $package.Workbook.Properties.Author = "Fabric Platform Discovery"
        $package.Workbook.View.ActiveTab = 0
        $sheetNames = @($package.Workbook.Worksheets | ForEach-Object Name)
        $package.SaveAs((New-Object IO.FileInfo($pending)))
        $package.Dispose()
        $package = $null
        $check = Open-ExcelPackage -Path $pending -ErrorAction Stop
        try {
            if ((($check.Workbook.Worksheets | ForEach-Object Name) -join "|") -cne ($sheetNames -join "|")) {
                throw "Saved workbook sheet verification failed."
            }
            foreach ($sheet in $check.Workbook.Worksheets) {
                foreach ($table in $sheet.Tables) {
                    foreach ($cell in $sheet.Cells[$table.Address.Start.Row,$table.Address.Start.Column,
                        $table.Address.Start.Row,$table.Address.End.Column]) {
                        if ($null -eq $cell.Comment) { throw "Saved workbook header note is missing." }
                    }
                }
            }
        } finally { $check.Dispose() }
        Move-Item -LiteralPath $pending -Destination $target -ErrorAction Stop
        [PSCustomObject]@{ Status = "Complete"; Path = $target; Sheets = $sheetNames; Error = "" }
    } finally {
        if ($null -ne $package) { $package.Dispose() }
        if (Test-Path -LiteralPath $pending) { Remove-Item -LiteralPath $pending -ErrorAction Stop }
        $script:AuditWorkbook = $null
    }
}
