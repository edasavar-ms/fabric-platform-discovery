# Project shared source data through the original workspace/principal renderer.
$script:WorkspaceColumns = @(
    "WorkspaceName","WorkspaceId","WorkspaceType","WorkspaceHostingMode",
    "CapacityId","CapacityName","CapacitySku","WorkspaceAdmins",
    "PrincipalDisplayName","PrincipalEmailOrUpn","PrincipalIdentifier",
    "PrincipalObjectId","PrincipalType","DirectoryUserType","WorkspaceRole",
    "State","IsOrphaned","IsReadOnly","IsOnDedicatedCapacity",
    "DefaultDatasetStorageFormat","HasWorkspaceLevelSettings","DashboardsCount",
    "ReportsCount","DatasetsCount","LargeModelCount","LargeModelSizeDataAvailable",
    "LargeStorageFormatCount","LargeStorageFormatModel","DataflowsCount",
    "FabricItemsCount","DiscoveryStatus","DiscoveryErrors"
)

function ConvertTo-AuditWorkspaceRows {
    param($Catalogue, $Scan, $FabricItems)
    $counts = @{}
    foreach ($item in $FabricItems.Items) {
        if ($NonFabricItemTypes -notcontains $item.type) {
            if (-not $counts.ContainsKey($item.workspaceId)) { $counts[$item.workspaceId] = 0 }
            $counts[$item.workspaceId]++
        }
    }
    $errorText = if ($FabricItems.Complete) { "" } else { "Fabric item inventory incomplete; see CollectionIssues.csv." }
    $rows = @(ConvertTo-DiscoveryWorkspaceRows -Workspaces @($Catalogue.Workspaces.Values) `
        -CapacityById $Catalogue.Capacities -ScanResults $Scan -FabricItemCounts $counts `
        -FabricItemInventoryAvailable $FabricItems.Complete -FabricItemInventoryError $errorText)
    foreach ($row in $rows) {
        if ($row.CapacityId -and -not $Catalogue.Capacities.ContainsKey($row.CapacityId)) {
            $row.DiscoveryStatus = "Partial"
            $row.DiscoveryErrors = @($row.DiscoveryErrors, "Capacity metadata unavailable." | Where-Object { $_ }) -join "; "
        }
        $row
    }
}

function Export-AuditWorkspaces {
    $catalogue = Get-AuditCatalogue
    $scan = Get-AuditScan
    $fabric = Get-AuditFabricItems
    $rows = @(ConvertTo-AuditWorkspaceRows $catalogue $scan $fabric)
    Write-AuditCsv "WorkspaceAccess.csv" $rows $script:WorkspaceColumns
    $columns = @($script:WorkspaceColumns | Where-Object {
        $_ -notmatch '^Principal' -and $_ -notin @("DirectoryUserType","WorkspaceRole")
    })
    $unique = @($rows | Group-Object WorkspaceId | ForEach-Object { $_.Group[0] })
    Write-AuditCsv "Workspaces.csv" $unique $columns
    if (@($rows | Where-Object DiscoveryStatus -eq "Partial").Count -gt 0) {
        Add-AuditIssue "Workspaces" "export" "IncompleteRows" "One or more workspaces have unavailable source data."
    }
}

# 2. Reconcile item identities, not names. App copies retain their own IDs.
function Get-AuditItemRows {
    $catalogue = Get-AuditCatalogue
    $fabric = Get-AuditFabricItems
    $scan = Get-AuditScan
    $items = @{}
    foreach ($item in $fabric.Items) {
        $wid = Get-AuditId $item.workspaceId
        $id = Get-AuditId $item.id
        $items["$wid/$id"] = [PSCustomObject][ordered]@{
            WorkspaceId = $wid
            WorkspaceName = [string](Get-AuditValue $catalogue.Workspaces[$wid] "name" "")
            CapacityId = [string](Get-AuditValue $catalogue.Workspaces[$wid] "capacityId" "")
            ItemId = $id; ItemName = [string](Get-AuditValue $item "name" (Get-AuditValue $item "displayName" ""))
            ItemType = [string](Get-AuditValue $item "type" "")
            FabricItemType = [string](Get-AuditValue $item "type" "")
            ScannerItemType = ""; Sources = "FabricAdminItems"
            AppId = ""; OriginalReportId = ""; DatasetId = ""
            CollectionStatus = $(if ($fabric.Complete) { "Complete" } else { "Partial" })
        }
    }
    $kinds = [ordered]@{
        reports = "Report"; dashboards = "Dashboard"; datasets = "SemanticModel"
        dataflows = "DataflowGen1"; datamarts = "Datamart"; lakehouses = "Lakehouse"
        warehouses = "Warehouse"; notebooks = "Notebook"; dataPipelines = "DataPipeline"
        eventhouses = "Eventhouse"; kqlDatabases = "KQLDatabase"; eventstreams = "Eventstream"
        environments = "Environment"; reflexes = "Reflex"; mlModels = "MLModel"
    }
    foreach ($wid in $scan.DetailsById.Keys) {
        if (-not $catalogue.Workspaces.ContainsKey($wid)) { continue }
        foreach ($kind in $kinds.Keys) {
            foreach ($item in @(Get-AuditValue $scan.DetailsById[$wid] $kind @())) {
                try {
                    $id = Get-AuditId (Get-AuditValue $item "id" (Get-AuditValue $item "objectId"))
                } catch {
                    Add-AuditIssue "Items" $wid "InvalidItemId" "$kind item omitted: missing or invalid ID."
                    continue
                }
                $key = "$wid/$id"
                if (-not $items.ContainsKey($key)) {
                    $items[$key] = [PSCustomObject][ordered]@{
                        WorkspaceId = $wid
                        WorkspaceName = [string](Get-AuditValue $catalogue.Workspaces[$wid] "name" "")
                        CapacityId = [string](Get-AuditValue $catalogue.Workspaces[$wid] "capacityId" "")
                        ItemId = $id; ItemName = [string](Get-AuditValue $item "name" (Get-AuditValue $item "displayName" ""))
                        ItemType = $kinds[$kind]; FabricItemType = ""; ScannerItemType = ""
                        Sources = ""; AppId = ""; OriginalReportId = ""; DatasetId = ""
                        CollectionStatus = "Complete"
                    }
                }
                $row = $items[$key]
                if ($row.ScannerItemType) {
                    Add-AuditIssue "Items" $key "DuplicateScannerItem" "Scanner returned this identity more than once."
                    $row.CollectionStatus = "Partial"
                    continue
                }
                $row.ScannerItemType = $kinds[$kind]
                $row.Sources = @($row.Sources, "WorkspaceScanner" | Where-Object { $_ }) -join "|"
                $row.AppId = [string](Get-AuditValue $item "appId" "")
                $row.OriginalReportId = [string](Get-AuditValue $item "originalReportObjectId" "")
                $row.DatasetId = [string](Get-AuditValue $item "datasetId" "")
                if ($scan.ErrorsById.ContainsKey($wid) -or -not $fabric.Complete) {
                    $row.CollectionStatus = "Partial"
                }
            }
        }
    }
    @($items.Values | Sort-Object WorkspaceId, ItemId)
}

function Export-AuditItems {
    $rows = @(Get-AuditItemRows)
    $columns = @("WorkspaceId","WorkspaceName","CapacityId","ItemId","ItemName","ItemType",
        "FabricItemType","ScannerItemType","Sources","AppId","OriginalReportId","DatasetId","CollectionStatus")
    Write-AuditCsv "Items.csv" $rows $columns
    $reconciliation = @($rows | Where-Object {
        -not $_.FabricItemType -or -not $_.ScannerItemType -or $_.FabricItemType -ne $_.ScannerItemType
    } | Select-Object WorkspaceId,ItemId,FabricItemType,ScannerItemType,Sources,AppId,OriginalReportId,CollectionStatus)
    Write-AuditCsv "ItemSourceReconciliation.csv" $reconciliation @(
        "WorkspaceId","ItemId","FabricItemType","ScannerItemType","Sources","AppId","OriginalReportId","CollectionStatus")
}
