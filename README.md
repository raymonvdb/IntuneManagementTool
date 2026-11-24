# Intune Management Tool

A read-only Windows tool with an Intune-like interface. For every policy, script, app or profile you see who it is assigned to, with include and exclude shown separately. The tool also offers a dashboard, an ASR overview, search by group or user, a details and settings view, and export. Nothing is changed in Intune.

## Starting

Double-click `Start-IntuneTool.bat`. The first time, you are asked whether the `Microsoft.Graph.Authentication` module may be installed. Then click **Sign in with Microsoft**. After signing in, the data is loaded automatically. Use **Load data** to refresh it and **Cancel** to interrupt a running action.

## The interface

The navigation is on the left, just like in Intune. The content is on the right.

| Group | Sections |
|---|---|
| Overview | Dashboard, Attention points, All assignments, Search by group/user |
| Devices | Configuration profiles, Compliance policies, Platform scripts, Remediations, Windows updates, Autopilot profiles, Enrollment, Devices |
| Endpoint security | All endpoint security, ASR rules, Attack surface reduction, Antivirus, Firewall, Disk encryption, EDR, Account protection, Security baselines |
| Apps | Apps, App protection policies, App configuration policies |
| Groups and filters | Groups in use, Assignment filters |

Each section shows its item count. In every list you can search, filter by platform and show only items that are not (fully) assigned. Sort columns by clicking the header. Double-click a row for details.

## What the tool shows

**Assignments.** Every row has a column with included groups (include) and a column with excluded groups (exclude). Assignment filters are shown as `[filter: Name (include)]`, and apps also show their intent (Required, Available, Uninstall). The status is *Assigned*, *Not assigned* or *Exclusion only*.

**Dashboard.** Cards with totals, unassigned items, groups in use, errors and ASR conflicts, plus a table per type. Click a card to jump to the matching list.

**Attention points.** Items that are not assigned or have only exclusions, references to groups that no longer exist, groups that are both included and excluded, and ASR rules with conflicting modes. The *Get member counts* button also detects groups without members.

**ASR rules.** Every Attack Surface Reduction rule per policy with its mode (Block, Audit, Warn, ...), include and exclude. If the same rule occurs in multiple policies, *Overlap* or *CONFLICT* is shown with the policies involved. The *Also scan Settings catalog* button also searches all Settings catalog profiles, because ASR is often configured there as well.

**Search by group/user.** Type the start of a name or UPN and pick a result. You see everything assigned to it, including via parent (nested) groups and via "All users/devices". The *Effect* column shows whether an exclusion wins over an inclusion. From *Groups in use*, a double-click takes you straight to this view.

**Details.** Double-click an item for tabs with assignments, properties and settings (Settings catalog with the real Intune names, administrative templates, endpoint security, OMA-URI and other properties). For platform scripts and remediations you see the script content. For a device you see the status of its configuration and compliance policies.

**Export.** Export any list to Excel (.xlsx, Excel does not need to be installed), CSV (semicolon) or an HTML report with search and sorting. From the dashboard you create a full report, in Excel with multiple sheets, including a flat overview with one row per assignment for pivot tables.

## Permissions

All permissions are read-only (delegated):

| Permission | Used for |
|---|---|
| `DeviceManagementConfiguration.Read.All` | profiles, compliance, scripts, endpoint security, filters, updates |
| `DeviceManagementApps.Read.All` | apps, app protection, app configuration |
| `DeviceManagementServiceConfig.Read.All` | Autopilot and enrollment |
| `DeviceManagementManagedDevices.Read.All` | devices and per-device policy status |
| `Group.Read.All` | group names |
| `GroupMember.Read.All` | member counts and memberships |
| `User.Read.All` | searching for users |

If your tenant requires admin consent for the Microsoft Graph PowerShell app, an administrator must approve it once. If a section is not available (for example because of a missing permission or license), the tool reports this after loading and the other sections keep working.

## Structure

| File | Content |
|---|---|
| `IntuneTool.ps1` | startup, building the navigation, wiring up events |
| `lib\Graph.ps1` | Graph calls with retry, all sources, groups and filters, loading |
| `lib\Settings.ps1` | reading settings, scripts and ASR rules |
| `lib\Analysis.ps1` | dashboard, attention points, group and filter overview, search |
| `lib\Export.ps1` | xlsx, csv and html export |
| `lib\Ui.ps1` | views, tables, dashboard, detail windows, signing in |
| `ui\MainWindow.xaml`, `ui\Resources.xaml` | layout and styles |

## Extending

Adding a new type is usually one block in `$script:Sources` in `lib\Graph.ps1` (endpoint, type, subtype, platform). A new navigation section is one line in `$script:Views` in `lib\Ui.ps1`.

## Good to know

- The tool uses the beta endpoints of Graph. These may change over time.
- Loading happens in the main window. The status bar shows progress and you can stop it with **Cancel**. It can take a while on large tenants.
- Assignments are retrieved in one go where possible with `$expand=assignments`. If that fails for a source, the tool falls back to one call per item.
- Group member counts are only retrieved on request, because this takes a separate call per group.
- Devices are only loaded when you open that page.
