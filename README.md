> [!IMPORTANT]
> **Disclaimer**
>
> This code is provided as a reference implementation to illustrate how the
> Fabric and Power BI administrative APIs can be used to produce a workspace and
> access discovery report. It is sample code. It is not a Microsoft product or
> service, is not a supported deliverable, and is not covered by any Microsoft
> support agreement, SLA, or warranty.
>
> It is provided "as is", without warranty of any kind, express or implied,
> including any implied warranty of merchantability or fitness for a particular
> purpose. The user or adopting organization is responsible for reviewing and
> testing the code before execution, for the privileges granted to the account
> that runs it, and for the handling, storage, and retention of its output, which
> contains workspace and identity information.
>
> Prepared by Eda Avar, Solution Engineer, Microsoft, as presales discovery
> input. For a production-grade, supported implementation, engage your Microsoft
> account team.
>
> This notice is also available in [DISCLAIMER.md](DISCLAIMER.md).

# Fabric Platform Discovery

An administrator-run PowerShell sample for Microsoft Fabric and Power BI
workspace, item, access and app discovery. One command creates a formatted
**Fabric Platform Discovery.xlsx** workbook and supporting CSVs.

The default run collects **Workspaces, Items, ReportAccess, AppAccess,
AppContent and ActivityEvents**, with **one completed UTC day of activity
events: yesterday UTC**. Inventory and access are current observations, not
historical snapshots. **Additional requirement 3 (historical CU consumption)
is DEFERRED and excluded from this release.** Activity events do not deliver
consumption, per-user CU or permissions history.

## Quick start

1. On this repository's GitHub page, choose **Code > Download ZIP** and extract
   the **entire package**, preserving the `audit` folder next to
   `Invoke-FabricDiscovery.ps1`. Alternatively, clone it:

   ```powershell
   git clone https://github.com/edasavar-ms/fabric-platform-discovery.git
   Set-Location .\fabric-platform-discovery
   ```

2. Open **Windows PowerShell 5.1 or PowerShell 7 on Windows** in the extracted
   folder. Install the dependencies once, in the **same PowerShell edition**
   you will use to run the script, subject to your organization's policy:

   ```powershell
   Install-Module Az.Accounts -Scope CurrentUser
   Install-Module ImportExcel -MinimumVersion 7.8.10 -Scope CurrentUser
   ```

3. Activate your **Fabric administrator** role if needed, then run:

   ```powershell
   .\Invoke-FabricDiscovery.ps1 -TenantId "<your tenant ID>" `
       -OutputDirectory "C:\Temp\Fabric Platform Discovery"
   ```

4. Complete the interactive Microsoft sign-in. Open the workbook at the
   printed path. **Review Summary and Collection Notes before interpreting
   any missing data or sharing results.**

Only `TenantId` is required; omit `OutputDirectory` to save under
`.\Fabric Platform Discovery`. No service principal, Key Vault, Python or
installed Excel is needed to generate the workbook. ISE and local Windows
administrator elevation are not required. Follow your organization's
execution-policy and signing requirements; the script never changes them
and never installs dependencies automatically.

## Before running

The signed-in account needs the relevant delegated admin-read access.
Classic published app content also requires the caller's actual app access;
the Fabric administrator role does not bypass it. See
[setup, permissions and parameters](docs/RUNNING.md).

Authentication uses `Az.Accounts` with `-SkipContextPopulation`, as in the
original discovery experience. **It still performs an initial read-only
Azure subscription metadata lookup.** This is separate from the Fabric/
Power BI collector allowlist; sign-in is not Entra-only. The collectors
exclude Graph, Azure resource inventory, permission changes, refreshes,
configuration and scheduling routes. The metadata scanner's POST is a
read-only metadata operation, not a workload refresh.

API reads are deliberately paced and retried within a **30-minute,
1,000-request default collection budget**. Even one day can contain many
activity pages. Large tenants may take longer and produce a bounded partial
run; review the manifest and increase the budget or select fewer collectors
only after agreeing the scope. Interactive sign-in and local workbook
generation are not forcibly interrupted by the collection deadline.

## What you receive

```text
Fabric Platform Discovery\
  <UTC timestamp>-<run ID>\
    Fabric Platform Discovery.xlsx
    Supporting files\
      RunManifest.json
      CollectionIssues.csv
      <selected CSV extracts>
```

The expected default workbook has **11 tabs**. **Workspace Access** is first
and preserves the original 32 headers, order and text. Every table header has
a description, repeated in **Column Guide**. Tables have filters, frozen
panes and clear coverage notes. Selective or failed runs can have fewer tabs.
Only this run's files are used: no fallback to previous runs.

The [workbook and requirements guide](docs/WORKBOOK-GUIDE.md) explains every
tab and what each requirement receives. Key boundaries:

- Individual report records remain **NeedsReview**, even if the person also
  has group/workspace access. This is **not proof of a direct-share violation**;
  the APIs do not expose grant origin. Groups are not flagged; applications
  are `NotEvaluated`. No group expansion or Graph lookup occurs.
- `WorkspaceAdmins` is an **administrator/contact proxy**, not the configured
  contact list. App-level principals are **not audience membership**.
- App content is **caller-visible, classic-app-only**. OrgApp metadata is
  retained separately; definitions needing item Write permissions are not read.
- `Partial` is legitimate when a supported read succeeds but a requested
  dimension is unavailable. Access denials and collection failures stay
  visible. A workbook's existence never proves a complete audit.

The launcher exits **0 for Complete** and **2 for Partial or Failed**, including
handled setup/output failures. PowerShell parameter-binding failures are
rejected before execution and can use the host's own nonzero exit code.
See [troubleshooting and optional commands](docs/RUNNING.md#troubleshooting).

Reports contain tenant and identity information. Keep them in an approved,
access-controlled location. Do not commit them to this public repository.
The workbook preserves formula-like names as literal text; raw CSV consumers
must import text columns explicitly rather than evaluate cell contents.

## Offline checks

The synthetic harness uses the real ImportExcel workbook engine, replaces
authentication/HTTP boundaries, and removes its temporary outputs:

```powershell
.\audit\checks\Test-FabricAudit.ps1
```

Run it in each installed PowerShell edition when changing the package. It
does not contact a tenant, sign in, or validate live permissions.
