# Workbook and requirements guide

Start with **Summary**, then **Collection Notes**. The expected default run
has the 11 tabs below when each collector produces its files; some may contain
no data. Selective runs or early failures can have fewer tabs. The first
**Workspace Access** table starts at row 1 with exactly the original 32
headers, in their original order and text. Other report tables start at row 4.
All headers have notes, also listed in **Column Guide**.

## Every default tab

| Tab (in order) | Grain and content | Requirement mapping / limits |
|---|---|---|
| Workspace Access | Workspace/direct role assignment. Workspaces without returned assignments still have one row with blank principal fields. Includes capacity/hosting, admin labels and inventory counts. | Original workspace/access requirement. 32-column contract retained; source limitations below apply. Do not sum repeated workspace counts across principals. |
| Summary | Run settings, selected collector statuses, requested UTC range and a map from every supporting CSV to its workbook representation. | Scope and completeness across all delivered requirements. A date range is not a historical inventory snapshot. |
| Workspaces | One row per active workspace; deduplicated workspace-level fields from Workspace Access. | Additional 1: all active workspaces returned across capacities, shared, personal and system types. Not deleted-workspace history. |
| Items | Workspace/item identity, name, type, current capacity and source labels. Reconciles Fabric Admin Items with scanner metadata by ID, not name. | Additional 2: actual items. Source-only records and type differences are retained; see reconciliation notes. Not inferred from counts alone. |
| Report Access | One returned report/principal access record with permission, source, finding and reason. | Additional 4: groups-only review. Filter `PolicyFinding = NeedsReview`; these are review candidates, not proven direct-share violations. |
| Apps | App identity/workspace/type and publication evidence. Classic apps and Fabric App/OrgApp/OrgAppAudience metadata stay distinct. | Additional 5, partial. Metadata existence does not prove publication; audience metadata is not a recipient list. |
| App Access | Returned classic app-level principals and their access rights. | Additional 5, partial. App users are not audience members; no audience-to-recipient mapping is delivered. |
| App Content | Workspace -> app -> caller-visible published report/dashboard identity. Keeps published and original report IDs distinct. | Additional 6, partial: classic only, not guaranteed tenant-complete. Workspace items are not substituted for published content. |
| Activity Events | Returned event ID, UTC timestamp, actor, action and available resource/correlation fields, one row per accepted event. | Supplementary history only, not any current-access or consumption requirement. Default: yesterday UTC only. |
| Collection Notes | Collection issues plus report, audience, published-content and activity coverage, and item-source reconciliation. Multiple filtered tables. | All requirements: denied, unsupported, partial and failed data remain explicit. |
| Column Guide | One row per workbook table header, with its sheet/section and description. | Data dictionary, including missing-value meaning, provenance and important interpretation limits. |

`ReportAccessFindings.csv` is an exact verified `NeedsReview` subset of
`ReportAccess.csv`, represented by the **Report Access filter**, not a duplicate
tab. Workbook generation checks that these rows agree.

## Requirement coverage

| Requirement | Delivered evidence | Boundary |
|---|---|---|
| Original workspace/principal/capacity/RBAC fields | Workspace Access retains the original 32-column output. Workspace names/types/hosting, capacity ID/name/SKU, admin/contact proxy, principal display name/email/identifier/object ID/type/directory user type and direct workspace role are included where the sources supply them. | Direct assignments, not effective access. Names are not unique; use IDs. Contact and directory limitations below apply. |
| Additional 1: all workspaces | Workspaces plus Workspace Access, across all returned active workspace types and hosting modes by default. | Optional capacity filter narrows output. No deleted history. |
| Additional 2: actual objects/items | Items and item-source reconciliation in Collection Notes. | Coverage follows both APIs; type/source differences remain visible rather than silently dropped. |
| Additional 3: historical CU consumption | **DEFERRED. Not included in this release.** | Activity events do **not** deliver CU consumption, per-user CU or permissions history. No CU collector, workbook tabs, model source setup or historical reconstruction is supplied. |
| Additional 4: direct individual report grants against a groups-only policy | Report Access / `NeedsReview` and its Reason; exact findings CSV retained. | **Partial assessment:** source schemas do not expose direct-versus-inherited grant origin. No confirmed direct-share violations are inferred. |
| Additional 5: apps, audiences and recipients | Apps and App Access; explicit coverage in Collection Notes. | **PARTIAL:** classic app principals are not audience membership. OrgApp metadata does not establish publication or audience recipients. |
| Additional 6: workspace -> app -> published content | App Content with app workspace, published ID, original report ID where returned and caller-visible coverage. | **PARTIAL:** classic supported content only; actual caller access required. OrgApp definitions needing item Write are excluded. |

## How to interpret the original workspace/access fields

The [full 32-column schema](OUTPUT-SCHEMA.md) is retained. `WorkspaceType`
comes from the API; `WorkspaceHostingMode` is derived from capacity assignment
and SKU. `DirectoryUserType` (for example Member/Guest, when supplied) is not
the workspace `Member` RBAC role. `PrincipalType` distinguishes people,
groups, applications, application profiles and entire-tenant entries.

`WorkspaceAdmins` lists direct Admin assignments as an **administrator/contact
proxy**, not the configured workspace contact list. Group principals remain
one assignment; their members are not expanded. Missing Entra IDs/user types
are not filled with Graph calls. `IsOrphaned` is only assessed for a successfully
scanned role-assignable workspace; missing scanner data, personal and system
workspaces do not become confirmed orphans.

Original size/inventory columns remain: `LargeModelCount` uses the 1 GiB
threshold only when model-size metadata is complete, and otherwise stays
blank. `LargeStorageFormatCount` / `LargeStorageFormatModel` describe
`PremiumFiles` settings, not proof of measured size.

## Groups-only report review

An individual report access record stays **NeedsReview**, even when the person
also has a workspace role or is believed to belong to an access group. That
does not establish whether the reported right was inherited or directly
granted. Review the actual report access separately before declaring a breach.
Entire-tenant entries also need broad-access review. Groups are
`GroupPrincipal` and are not flagged by this check; approval of a particular
group is outside scope. Applications/profiles and unknown types are
`NotEvaluated`. There is no directory or group-membership expansion.

## Missing data, errors and output safety

`Complete` means a supported read completed, not that every requested business
dimension is exposed by the API. `Partial` can be expected for report grant
origin, app audience membership or caller-visible content. `Failed` and
`Unsupported` are not successful empty results. In particular, a classic app
can be listed while its reports or dashboards are denied; that error remains
in Collection Notes. The script neither hides it nor changes access.

Activity coverage is per UTC day. Deduplicated events and missing capacity
fields under a capacity filter are counted. No capacity is guessed from
today's inventory. Completed paging is not a promise that upstream audit
recording was complete, instantaneous or retained for every operation.

Only this run's CSVs are read into the workbook. IDs and formula-like names
stay literal text. Numeric fields retain zeros; nulls remain blank; values
beyond Excel's reliable precision remain text. The first sheet preserves
all source values as text for compatibility. Raw CSVs also preserve source
text, so import them as text rather than opening untrusted formula-like
strings through automatic spreadsheet conversion.

If workbook generation fails, the manifest records the failure and available
CSVs remain for inspection. A storage failure can also prevent saving the
manifest; the launcher then prints a failure rather than claiming success.
Handle all outputs under your tenant-data retention and access policies.
