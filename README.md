# MDM On-Demand Actions (MDM-ODA) &#x26A1;

**Live Analytics, Insights & Actions for Entra ID and Intune**

[![Version](https://img.shields.io/badge/Version-0.66-green)](https://github.com/satishsinghi-gh/mdm-oda/releases)
[![PowerShell 7](https://img.shields.io/badge/PowerShell-7.x-blue?logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![WPF](https://img.shields.io/badge/UI-WPF-blueviolet)](https://learn.microsoft.com/en-us/dotnet/desktop/wpf/)
[![Microsoft Graph](https://img.shields.io/badge/API-Microsoft%20Graph-0078D4?logo=microsoft)](https://learn.microsoft.com/en-us/graph/)
[![Blog](https://img.shields.io/badge/Blog-GitHub%20Pages-green)](https://satishsinghi-gh.github.io/MDM-ODA/)

---

![MDM-ODA Group Management Blades Overview](blades-overview.png)

## Overview

MDM-ODA is a PowerShell & WPF based plug-n-play tool for Entra & Intune on-demand operations. Built with attention to detail for the granular challenges faced by support teams, enabling project teams to get reliable, meaningful, up-to-date insights and reports on the go. Built with safeguards to prevent accidental actions, keeping Zero Trust and least privilege as top priority.

> For a full deep-dive into the tool's design, architecture, and security model, visit the [MDM-ODA Blog](https://satishsinghi-gh.github.io/MDM-ODA/).

### Project Mission

- Maximize operational efficiency
- Automate variety of on-demand actions that need to be performed on the go
- Deliver a thoughtful, well-crafted experience that IT professionals genuinely enjoy using
- Bring together data from different portals/pages using a single page — no browser tab madness
- Surface actionable insights directly — no Excel exports, no manual pivot tables, just immediate clarity for faster triage
- Automate bulk actions on-demand that are natively not possible
- Reduce human errors using validation workflows
- Save hours of efforts and endless fatigue caused by repetitive tasks
- No manual setup needed, no admin rights needed — just plug and play
- Make Click-Ops great again

## Highlights

| | |
|---|---|
| **Lightweight & Powerful** | Enterprise-grade functionality built entirely on native Windows components — PowerShell 7 and WPF. Zero third-party dependencies, zero licensing. All data is live from Microsoft Graph — no Power BI refresh cycles, no stale dashboards. |
| **Security & Guardrails** | Delegated auth flow for least privilege — use your Tenant & Client ID. Validation & preview before each write action. Live verbose logging for transparency. No admin rights required. |
| **Productivity** | Complex on-demand actions with minimum effort. In-page live table output with flexibility to select cells, copy individual cell/row/all, or export to XLSX. |

## Features — Group Management

<details>
<summary><strong>Search Entra Objects</strong> — Keyword search across multiple Entra object types from a single page</summary>

![Search Entra Objects](search-entra-objects.png)
*Search across Users, Groups, Devices, and Service Principals with real-time filtering*
</details>

<details>
<summary><strong>List Group Members</strong> — List members from multiple groups with a single click</summary>

![List Group Members](list-group-members.png)
*Query members from multiple groups simultaneously with comprehensive details*
</details>

<details>
<summary><strong>Object Membership</strong> — Find group membership for bulk items (Users/Devices/Groups)</summary>

![Object Membership](object-membership.png)
*Bulk lookup of group membership across users, devices, and groups*
</details>

<details>
<summary><strong>Find Groups by Owners</strong> — Enter UPNs, get all groups they own</summary>

![Find Groups by Owners](find-groups-by-owners.png)
*Identify all groups owned by specific users with detailed ownership insights*
</details>

<details>
<summary><strong>Create Group</strong> — Create Security/M365 Group with bulk owners, members, or dynamic query from a single page</summary>

![Create Group](create-group.png)
*Streamlined group creation with owners, members, and dynamic rules — no CSV, no browser navigation*
</details>

<details>
<summary><strong>Set Bulk Owners on Bulk Groups</strong> — Assign multiple owners to multiple groups in one operation</summary>

![Set Bulk Owners](set-bulk-owners.png)
*Bulk owner assignment across multiple groups with validation*
</details>

<details>
<summary><strong>Add User Devices to Groups</strong> — Enter UPNs, auto-resolve their registered devices and add to groups</summary>

![Add User Devices to Groups](add-user-devices.png)
*Device-to-group assignment driven by user identity*
</details>

<details>
<summary><strong>Find Common/Distinct Groups</strong> — Compare group memberships across multiple objects</summary>

![Find Common/Distinct Groups](find-common-distinct.png)
*Identify overlapping and unique group memberships for users, devices, or groups*
</details>

<details>
<summary><strong>Compare Groups</strong> — Side-by-side comparison of group properties and memberships</summary>

![Compare Groups](compare-groups.png)
*Detailed group comparison with property and membership diff*
</details>

### Additional Group Management

- **Rename Bulk Groups** — Rename multiple groups at once
- **Update Dynamic Membership Rules** — Modify dynamic queries on existing groups
- **Delete Empty Groups** — Safely remove groups with zero members (with confirmation)

## Features — Device Management

<details>
<summary><strong>Device Info</strong> — Comprehensive device details from Entra and Intune in one view</summary>

![Device Info](device-info.png)
*Hardware, OS, compliance, encryption, and registration details from a single query*
</details>

<details>
<summary><strong>Intune Policy Assignments</strong> — View all policies assigned to a device through its group memberships</summary>

![Intune Policy Assignments - Overview](policy-assignments-1.png)
*Policy assignment overview with group context*

![Intune Policy Assignments - Details](policy-assignments-2.png)
*Detailed policy breakdown with assignment intent and filter evaluation*

![Intune Policy Assignments - Expanded](policy-assignments-3.png)
*Full policy assignment landscape across configuration profiles, compliance, and apps*
</details>

## Productivity Features

<details>
<summary><strong>Session Notes</strong> — Built-in notepad for each session with timestamp and context</summary>

![Session Notes](session-notes.png)
*Take notes during operations without leaving the tool*
</details>

<details>
<summary><strong>Verbose Logging & Keyword Filter</strong> — Real-time operation logging with search</summary>

![Keyword Filter](verbose-logging.png)
*Filter logs by keyword for quick troubleshooting*

![Verbose Logging](prereq-handling-2.png)
*Detailed operation logs with timestamps*
</details>

<details>
<summary><strong>Prerequisite Handling</strong> — Automatic detection and installation of dependencies</summary>

![Prerequisite Check](prereq-handling-1.png)
*Automatic detection of system prerequisites*

![Prerequisite Installation](prereq-handling-2.png)
*Installation progress and status reporting*
</details>

### Additional Productivity Controls

- **Clear Inputs** — Clear the page and start fresh with a single click
- **Stop Operation** — Cancel ongoing operations at any time without waiting for completion
- **Feedback** — Built-in feedback mechanism to report issues or suggest improvements

## Security & Auth Design

MDM-ODA uses the OAuth 2.0 delegated flow exclusively — the app never holds standalone permissions. Every API call executes in the context of the signed-in user, meaning the effective permission is always the intersection of what the app registration allows and what the user's Entra/Intune roles permit. The recommended configuration uses **read-only API scopes** for everyday operations. Write permissions are only needed when performing create, update, or delete operations.

Every write action follows a strict validation-before-commit workflow: the tool validates input format, checks for duplicates, resolves object identifiers, and presents a structured preview of pending changes. Only after the user explicitly confirms does the operation execute.

> For the full architecture diagram and detailed auth flow, see the [blog](https://satishsinghi-gh.github.io/MDM-ODA/).

## Prerequisites

1. **Windows 11** with WPF (built-in, no additional installation needed)
2. **PowerShell 7** — handled automatically by the script (auto-installs via winget if missing). The orchestrator can be launched from a standard PowerShell 5.1 host — it detects the running version, locates or installs PS7, and re-launches itself in the PS7 runtime automatically
3. **Internet Connectivity** — required for PowerShell Gallery modules and Microsoft Graph API
4. **No Admin Rights Required** — MDM-ODA runs in user context
5. **Code Signing & WDAC** — if WDAC or script execution policies are enforced, code signing adjustments may be needed

## Getting Started

```powershell
# 1. Clone the repository
git clone https://github.com/satishsinghi-gh/mdm-oda.git

# 2. Configure credentials (optional) — pre-populate your Tenant ID and Client ID
#    in the script, or enter them manually at launch

# 3. Launch the downloaded script — no parameters, no admin rights needed
.\MDM-ODA.ps1

# 4. Authenticate with your Entra credentials and start using the tool
```

### What's Included

- Fully functional PowerShell 7 script with embedded WPF UI
- Automatic prerequisite detection and installation
- Microsoft Graph SDK integration for reliable API calls
- Real-time verbose logging to local file system
- Validation workflows for write operations
- Export to Excel (XLSX) capability
- Complete source code and documentation

## Permissions

### Delegated App Permissions

| Permission | Purpose |
|---|---|
| `User.Read` | Sign-in and read current user profile (/me for PIM checks) |
| `User.Read.All` | Resolve UPN inputs and read user properties across all functions |
| `Group.Read.All` | Read group properties, list groups, read types and membership rules |
| `GroupMember.Read.All` | List group members and query member counts |
| `Directory.Read.All` | TransitiveMemberOf for PIM role detection and object membership |
| `Device.Read.All` | Resolve devices, read properties, query registered users |
| `DeviceManagementConfiguration.Read.All` | Read Intune config profiles and policies for assignment lookups |
| `DeviceManagementManagedDevices.Read.All` | Query managed devices by Azure AD device ID or serial number |
| `DeviceManagementRBAC.Read.All` | Read Intune RBAC settings (role assignments and role definitions) |
| `offline_access` | Maintain refresh token for persistent session |

> **Note:** The documented least-privileged permissions for group write operations are `Group.ReadWrite.All` and `GroupMember.ReadWrite.All`. However, based on testing, group owners with scoped Intune RBAC roles can perform all write operations with only the read-only scopes above. If you want to guarantee write access regardless of ownership, add `Group.ReadWrite.All` and `GroupMember.ReadWrite.All`.

### User Permissions

Entra built-in roles or custom RBAC roles determine which specific resources a user can access. The app permissions set the API surface ceiling, but Intune RBAC and group ownership scope the actual access. **Group Owners** is sufficient for most group management operations. For comprehensive device and policy insights, users may benefit from **Intune Reader** or **Intune Administrator** roles depending on scope.

### Web Application Redirect URI

For WAM (Web Account Manager) based authentication, configure the following redirect URI in your Entra app registration:

```
ms-appx-web://Microsoft.AAD.BrokerPlugin/{Client-ID}
```

Replace `{Client-ID}` with your actual Application (client) ID from Entra.

## Roadmap

MDM-ODA is actively evolving. Here's what's planned for upcoming releases:

- **Input-Based Bulk Actions** — Sync, Remediation (excluding destructive actions like Wipe/Delete)
- **Comprehensive Update Insights** — Quality Updates, Feature Updates, Driver Updates
- **Application Landscape** — Platforms, Assignments, Deployment States, Filters, App Creation Workflows
- **Defender Integration** — Timeline Events, Advanced Hunting, Vulnerability State, Software Inventory
- **Advanced Policy Management Actions** — Targeted modifications, cloning, bulk assignment management
- **Advanced Dynamic Group Query Builder** — Visual query builder with syntax validation and preview
- **Log Analytics Integration** — Extended Hardware Inventory and Audit data from Log Analytics

## Author

**Satish Singhi**

> [!WARNING]
> **Disclaimer:** This tool is provided "as-is" without warranty of any kind, express or implied. The author assumes no liability for any damages arising from its use. Always validate operations in a non-production environment before deploying to production tenants.

## License

This project is licensed under the [MIT License](LICENSE).
