<#
.SYNOPSIS
Creates a Microsoft Fabric inventory, access and activity discovery workbook.
.DESCRIPTION
All six supported extracts run by default. A formatted Excel workbook is created automatically
in a unique run directory. Supporting CSVs and collection status are retained
in its Supporting files subfolder. Requires Az.Accounts and ImportExcel 7.8.10+;
Python and Excel are not required. Collector requests exclude Graph, Azure resource,
refresh, configuration, permission-change, and scheduling endpoints.

Authentication uses the same Az.Accounts approach as v1 and can read Azure
subscription metadata despite SkipContextPopulation. This sign-in behavior
must be accepted separately from the Fabric/Power BI collector scope.

Activity history defaults to one completed UTC day: yesterday UTC.
Inventory and access are current observations, not historical permissions.
See README.md and docs\WORKBOOK-GUIDE.md for prerequisites and source limitations.
Sample code provided as is, without warranty or Microsoft support.
Prepared by Eda Avar, Solution Engineer, Microsoft, as presales discovery input.
.EXAMPLE
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" `
    -OutputDirectory "C:\Temp\Fabric Platform Discovery"
.EXAMPLE
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" -Collect Workspaces,Items
#>
#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TenantId,
    [Parameter()][ValidateNotNullOrEmpty()][string]$OutputDirectory = (Join-Path (Get-Location) "Fabric Platform Discovery"),
    [Parameter()]
    [ValidateSet("Workspaces", "Items", "ReportAccess", "AppAccess", "AppContent", "ActivityEvents")]
    [string[]]$Collect = @("Workspaces", "Items", "ReportAccess", "AppAccess", "AppContent", "ActivityEvents"),
    [Parameter()][Guid[]]$CapacityId = @(),
    [Parameter()][ValidateRange(1, 28)][int]$HistoryDays = 1,
    [Parameter()][ValidatePattern('^\d{4}-\d{2}-\d{2}$')][string]$EndDateUtc,
    [Parameter()][ValidateRange(1, 5000)][int]$WorkspacePageSize = 500,
    [Parameter()][ValidateRange(1, 100)][int]$ScanBatchSize = 100,
    [Parameter()][ValidateRange(30, 3600)][int]$ScanTimeoutSeconds = 600,
    [Parameter()][ValidateRange(0, 10)][int]$MaxRetries = 5,
    [Parameter()][ValidateRange(1, 10000)][int]$MaxRequests = 1000,
    [Parameter()][ValidateRange(1, 120)][int]$MaxRunMinutes = 30
)

try {
    Import-Module (Join-Path $PSScriptRoot "audit\FabricAudit.psm1") -Force -ErrorAction Stop
    $result = Invoke-FabricAuditRun @PSBoundParameters
    if ($result.WorkbookPath) { Write-Host ("Workbook: {0}" -f $result.WorkbookPath) }
    Write-Host ("Discovery status: {0}. Manifest: {1}" -f $result.Status, $result.ManifestPath)
    if ($result.Status -ne "Complete") {
        Write-Warning "Collection or workbook output is incomplete. Review Summary, Collection Notes and RunManifest.json; missing data is not an empty successful result."
        exit 2
    }
} catch {
    Write-Error ("Discovery failed: {0}" -f $_.Exception.Message) -ErrorAction Continue
    exit 2
}
