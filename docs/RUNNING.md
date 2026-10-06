# Setup and running

## Software and permissions

Use Windows PowerShell 5.1 or PowerShell 7 on Windows. Install `Az.Accounts`
and `ImportExcel` **7.8.10 or later** in that edition, using the quick-start
commands in [README](../README.md). ImportExcel is a third-party dependency;
review it under your organization's policy. No Excel installation is needed
for generation. An Excel-compatible viewer is needed to open the result.

Both dependencies and the column-definition file are checked before sign-in.
Extract the whole package; do not copy just the launcher or paste fragments.
Local administrator elevation and PowerShell ISE are not required. Do not
bypass managed execution policies; obtain approved signing/unblocking guidance
from your organization if execution is restricted.

| Source | Existing access needed |
|---|---|
| Workspace/capacity catalogue, scanner, report users, classic app users | An interactive Microsoft Entra user with an active Fabric administrator role and authorization for the corresponding delegated admin-read APIs. Activate PIM before running. |
| Fabric Admin Items | Authorization to use the Fabric read-only Admin Items API. Source coverage can differ from the workspace scanner. |
| Classic published app reports and dashboards | Relevant delegated app/content-read permissions and actual access to that published app. Being tenant administrator alone does not establish visibility. |
| Activity Events | Delegated admin-read access and a requested range within the last 28 completed UTC days. Retention and source availability still apply. |

The package does not grant consent, activate roles, change tenant settings or
elevate app permissions. A denied read stays denied and is recorded.
OrgApp/OrgAppAudience definitions requiring item Write are excluded.

Allow HTTPS access to Microsoft Entra sign-in endpoints (including
`login.microsoftonline.com` and those required by conditional-access policy),
`api.powerbi.com` and `api.fabric.microsoft.com`. `Az.Accounts` also performs
its initial read-only Azure subscription metadata lookup during sign-in,
even with `-SkipContextPopulation`. Disabling context autosave limits
process context persistence; it does not remove that lookup or all SDK token
caching. Accept this authentication behavior separately from collection.

## Optional parameters

| Parameter | Default | Meaning |
|---|---|---|
| `TenantId` | Required | Your tenant GUID or verified tenant domain. |
| `OutputDirectory` | `.\Fabric Platform Discovery` | Parent of a new timestamp/GUID directory. Existing runs are not overwritten. |
| `Collect` | `Workspaces,Items,ReportAccess,AppAccess,AppContent,ActivityEvents` | One or more of these six collectors. Shared inventory is reused within a run. |
| `CapacityId` | No filter | Capacity GUID array. Filters inventory/access by current placement and activities by the event's own capacity field. |
| `HistoryDays` | `1` | 1-28 completed UTC days, inclusive; applies only to ActivityEvents, not inventory or access. |
| `EndDateUtc` | Yesterday UTC | Inclusive end date in `yyyy-MM-dd`. Today/future dates are rejected. The entire activity range must fit the 28-day lookback. |
| `WorkspacePageSize` | `500` | 1-5,000 catalogue rows per request. Smaller pages may help response size but cost more rate-limited calls. |
| `ScanBatchSize` | `100` | 1-100 workspaces per scanner batch. Smaller batches may help large metadata responses. |
| `ScanTimeoutSeconds` | `600` | 30-3,600 seconds per scanner batch, also subject to the overall collection deadline. |
| `MaxRetries` | `5` | 0-10 transient retries per request; one additional token refresh is allowed after HTTP 401. |
| `MaxRequests` | `1000` | 1-10,000 HTTP attempts, including retries and scanner polls. |
| `MaxRunMinutes` | `30` | 1-120 minutes for sign-in/collection budget. Guards requests and waits, not an OS-level watchdog. Local packaging follows collection. |

**UTC is not the computer's local calendar day.** The default uses
`[DateTime]::UtcNow.Date.AddDays(-1)`, covering 00:00:00.000 through
23:59:59.999 UTC of that day. `HistoryDays 7` means seven completed UTC days
ending yesterday, not the last 168 hours. Inventory/access are sequential
current reads; historical dates do not reconstruct them.

Inventory-only run:

```powershell
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" -Collect Workspaces,Items
```

Seven completed days of activity history, without inventory/access:

```powershell
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" `
    -Collect ActivityEvents -HistoryDays 7
```

Explicit completed UTC end day (still inside the supported lookback):

```powershell
$endDay = [DateTime]::UtcNow.Date.AddDays(-2).ToString("yyyy-MM-dd")
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" `
    -Collect ActivityEvents -HistoryDays 1 -EndDateUtc $endDay
```

Capacity-filtered output with a larger, agreed collection budget:

```powershell
.\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" `
    -CapacityId "<your capacity GUID>" -MaxRunMinutes 60 -MaxRequests 2000
```

**A capacity filter is not an API isolation boundary.** The script reads tenant
catalogues and Fabric item pages before filtering. App collection reads the
tenant classic-app catalogue. Scanner submissions include selected workspaces
only. Activity collection reads full UTC-day pages, then filters by the
event's own `CapacityId`; missing capacity fields are counted as excluded,
not filled from today's workspace placement. Shared/personal workspaces
without a matching capacity are excluded by a capacity filter.

## Read-only collection boundary

Only these routes are admitted by the collector transport:

| Service | Method and path |
|---|---|
| Power BI | GET `/v1.0/myorg/admin/capacities`, `/admin/groups`, `/admin/apps`, `/admin/activityevents` (all under `/v1.0/myorg`) |
| Power BI metadata scanner | POST `/v1.0/myorg/admin/workspaces/getInfo?getArtifactUsers=true`; GET `/v1.0/myorg/admin/workspaces/scanStatus/{id}` and `/scanResult/{id}` |
| Power BI access lists | GET `/v1.0/myorg/admin/reports/{id}/users` and `/v1.0/myorg/admin/apps/{id}/users` |
| Power BI classic published content | GET `/v1.0/myorg/apps/{id}/reports` and `/v1.0/myorg/apps/{id}/dashboards` |
| Fabric inventory | GET `/v1/admin/items` |

The scanner POST starts a **read-only metadata scan**, not a refresh or
configuration change. The transport rejects Graph, ARM/resource inventory,
refresh, grants, configuration, other methods and cross-host redirects.
Continuation tokens are kept on the approved service host.
No data-model queries, ongoing capture, scheduling or backfill from old runs
are performed. Source APIs, including preview capabilities, can change.

Reads with known hourly limits are paced: workspace pages roughly every
72 seconds, scan submissions roughly every 7 seconds, and app/user/activity/
Fabric item reads roughly every 18 seconds within their respective buckets.
Other tenant traffic can still cause 429 throttling. One completed day is a
small *time range*, not necessarily a small number of records or requests.
Bounded retries honor Retry-After when available.

The deadline prevents new requests or waits beyond the budget. It cannot
forcibly interrupt interactive sign-in, stop a server-side scan already
submitted, or bound local file-writing time. Completed data and explicit
failure/coverage notes remain available when feasible. There is no automatic
resume or merging of separate runs.

## Troubleshooting

| Symptom | Meaning and action |
|---|---|
| Dependency failure before sign-in | Install/update the named dependency in the edition you are running. Check repository policy, module paths and package extraction. No automatic install is attempted. |
| Script blocked by execution policy | Follow organizational policy for approved signing/unblocking. Do not change managed policy merely to run this sample. |
| Sign-in or HTTP 401/403 | Confirm tenant, active role, delegated authorization and conditional-access requirements. For app content, confirm actual published-app access. Do not relax permissions just to make status Complete. |
| `Partial` with usable workbook | Review Summary and Collection Notes. Unsupported grant origin, audience membership and content visibility legitimately make the default run Partial, even if all supported reads succeed. |
| `Failed` or no workbook | Read printed errors, `RunManifest.json` and available supporting CSVs. If workbook creation fails, its path remains blank and its error is recorded. Correct the output/dependency issue and run again into a new directory. |
| Write/permission/disk error | Choose a writable, access-controlled path with space. CSV or workbook errors never become successful empty results. If even the manifest cannot be saved, rely on the console error; no durable result is guaranteed. |
| Budget reached or repeated 429 | Agree a larger budget, fewer collectors, or reduced output scope. Check coverage before relying on results; changing page size has tradeoffs. |
| Blank field or zero rows | Blank can mean unknown, not applicable, denied or not exposed. Zero is meaningful only for a successful supported read. Coverage counters with Failed/Unsupported do not prove there is no data. |
| Fewer workbook tabs | Only files created in this run are represented. Unselected or failed-before-export collectors do not get misleading placeholder report tabs. |
| Workbook size/cell limit | Excel has row and cell-length limits. The exporter fails visibly instead of silently truncating; supporting CSVs are retained where written. |

Exit **0** means all selected collectors and workbook creation completed
within their supported scope. Exit **2** means Partial/Failed, including
handled prerequisite and output failures. Invalid arguments are rejected by
PowerShell before the script body and may use its own nonzero code.

Protect the workbook, manifest, warnings and CSVs as sensitive tenant data.
Header descriptions are available in [Column Guide](WORKBOOK-GUIDE.md) and
the [original schema reference](OUTPUT-SCHEMA.md). Do not publish live outputs
or authentication logs in repository issues.
