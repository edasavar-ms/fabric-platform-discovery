# Output schema

`WorkspaceAccess.csv` and the first **Workspace Access** worksheet emit these
original 32 columns in this exact order. Worksheet values remain source text.
Additional tabs are described in [WORKBOOK-GUIDE.md](WORKBOOK-GUIDE.md).

| Column | Description |
|---|---|
| `WorkspaceName` | Workspace display name. |
| `WorkspaceId` | Workspace GUID. |
| `WorkspaceType` | Workspace type returned by the Power BI Admin API. |
| `WorkspaceHostingMode` | Derived hosting classification: Fabric, Premium, Embedded, PPU, other dedicated, or shared capacity. |
| `CapacityId` | Assigned dedicated capacity GUID; blank for shared capacity. |
| `CapacityName` | Capacity display name; blank for shared capacity. |
| `CapacitySku` | Capacity SKU; blank for shared capacity. |
| `WorkspaceAdmins` | Semicolon-separated direct Admin assignments: an administrator/contact proxy, not the configured contact list or expanded group members. |
| `PrincipalDisplayName` | Display name of the directly assigned principal. |
| `PrincipalEmailOrUpn` | Email/UPN where available; otherwise the API identifier. |
| `PrincipalIdentifier` | Identifier returned by the Power BI Admin API. |
| `PrincipalObjectId` | Microsoft Entra object ID where available. |
| `PrincipalType` | Friendly type: Individual, AD Group, Service Principal, Service Principal Profile, or Entire Tenant. |
| `DirectoryUserType` | Source directory user type, such as Member or Guest, when returned. |
| `WorkspaceRole` | Direct workspace role: Admin, Member, Contributor, or Viewer. |
| `State` | Workspace state. The default report includes only Active workspaces. |
| `IsOrphaned` | True when scanner data confirms an active role-assignable workspace has no direct Admin assignment. Blank for personal/system workspaces or unavailable scanner data. |
| `IsReadOnly` | Existing discovery field returned by the workspace API. |
| `IsOnDedicatedCapacity` | Existing discovery field indicating dedicated-capacity assignment. |
| `DefaultDatasetStorageFormat` | Existing workspace semantic-model storage setting. |
| `HasWorkspaceLevelSettings` | Existing workspace-level settings indicator. |
| `DashboardsCount` | Existing count of Power BI dashboards. Blank when workspace scanner data is unavailable. |
| `ReportsCount` | Existing count of Power BI reports. Blank when workspace scanner data is unavailable. |
| `DatasetsCount` | Existing count of Power BI semantic models/datasets. Blank when workspace scanner data is unavailable. |
| `LargeModelCount` | Existing count of datasets at least 1 GiB. Blank when the API omits size metadata. |
| `LargeModelSizeDataAvailable` | True when `LargeModelCount` is based on complete size metadata, including workspaces with no datasets. |
| `LargeStorageFormatCount` | Existing count of datasets using `PremiumFiles`; blank when scanner data is unavailable. |
| `LargeStorageFormatModel` | Existing pipe-separated names of datasets using `PremiumFiles`. |
| `DataflowsCount` | Existing count of Power BI dataflows; blank when scanner data is unavailable. |
| `FabricItemsCount` | Count of Fabric-native workloads returned by the Fabric Admin API. Power BI reports, semantic models, dashboards, apps, and datamarts are excluded; Dataflow Gen2 and newly introduced Fabric item types are included. |
| `DiscoveryStatus` | `Complete` or `Partial`. |
| `DiscoveryErrors` | Explicit reason for partial discovery; blank for a complete row. |

## Grain

The report grain is one row per workspace and direct role assignment. An active
workspace with no direct role assignments still receives one row with blank
principal fields so that no workspace disappears from the report.

Group membership is not expanded. An Entra group assigned to three workspaces
produces one group row in each workspace, not one row for every group member.

Workspace inventory and RBAC assignments come from the supported Workspace
Scanner API in batches of no more than 100 workspaces. If a scan fails or omits
a workspace, its row is marked `Partial`, its scanner-derived fields are left
blank, and the error is recorded. Missing RBAC data is never reported as a
confirmed orphaned workspace.
