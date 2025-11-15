# Maester – Protection Sets Run

This repository contains two PowerShell scripts designed to assess a Microsoft 365 tenant with **Maester** and export results aligned to **DocuForge Mapper V3**.

- **`MaesterInteractive.ps1`** – Runs the Maester tests for a specific customer (defaults to `ProtectionSets`) and exports results to the Mapper folder.
- **`Install-MaesterPrereqs-AllUsers.ps1`** – One‑time, **AllUsers** prerequisites installer that prepares a machine to run the main script; ends with a popup listing recommended admin roles.

> **Author:** Yoni Meeus  
> **Last updated:** 2025-11-15

---

## 1) File Overview

### 1.1 MaesterInteractive.ps1
**Purpose:**
- Installs/loads the **Maester** module and required service modules (Graph, Exchange Online; Teams/SPO optional).
- Connects interactively to Microsoft Graph (and optionally EXO/Teams/SPO).
- Runs Maester tests, saving outputs to the customer’s folder in a structure DocuForge Mapper V3 can consume.
- Produces:
  - `ReportMapper/maester_failed_detailed.csv`
  - `ReportMapper/maester_skipped_detailed.csv`
  - `ReportMapper/MaesterResults.html` (reference)
  - `ReportMapper/MaesterResults.json` (severity enrichment)

**Key defaults & behavior:**
- **Customer name default:** `ProtectionSets`. You can override with `-CustomerName "<Name>"`.
- **Paths:** `C:\M365Factory\Customers\<CustomerName>\ReportMapper`.
- **Connections:** Graph is always used; EXO on by default; Teams/SPO are optional toggles in the script variables.
- **Severities:** Normalizes severity using test metadata, JSON enrichment, and tag/title heuristics with a safe fallback to `Informational`.
- **Finish:** Shows a completion message box and then exits.

### 1.2 Install-MaesterPrereqs-AllUsers.ps1
**Purpose:**
- Prepares any Windows machine for running `MaesterInteractive.ps1` without prompts or missing modules.

**What it does:**
- Ensures TLS 1.2, NuGet provider, and trusted PowerShell Gallery.
- Installs/updates modules at **AllUsers** scope:
  - `Pester (>= 5.5.0)`
  - `Microsoft.Graph (>= 2.0.0)`
  - `ExchangeOnlineManagement (>= 3.0.0)`
  - `Maester (>= 1.0.0)`
  - *(Optional)* `MicrosoftTeams (>= 5.0.0)` and `PnP.PowerShell (>= 2.1.0)`
- Validates by importing modules and checking key commands.
- Ends with a **popup** listing **recommended admin roles** for the widest read/assessment coverage (can be suppressed with `-NoRolesPopup`).

---

## 2) Prerequisites

- **Windows PowerShell** (5.1+) in an **elevated** session for AllUsers installs.
- Internet access to `https://www.powershellgallery.com/` (or an internal mirror) to download modules.
- The account used at runtime must have sufficient **read permissions**. See [Permissions & Roles](#5-permissions--roles) below.

---

## 3) Installation & First Run

### 3.1 Install prerequisites (run once per machine)
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
# Add optional components as needed
./Install-MaesterPrereqs-AllUsers.ps1 -IncludeTeams -IncludeSPO
```

### 3.2 Run the Maester assessment
```powershell
# Default customer is "ProtectionSets"
./MaesterInteractive.ps1

# Or override the customer name
./MaesterInteractive.ps1 -CustomerName "Contoso"
```

---

## 4) Output Locations

Given a customer name **`<CustomerName>`**, outputs are placed at:
```
C:\M365Factory\Customers\<CustomerName>\ReportMapper\
  ├─ maester_failed_detailed.csv
  ├─ maester_skipped_detailed.csv
  ├─ MaesterResults.html
  └─ MaesterResults.json
```
These are aligned with **DocuForge Mapper V3** expectations.

---

## 5) Permissions & Roles

At the end of the prerequisites script, a popup summarizes the recommended roles. For convenience:

- **Tenant‑wide (Entra ID / M365)**
  - Global Reader
  - Security Reader
  - Reports Reader
- **Conditional Access / Identity**
  - Security Reader *(often sufficient for read)* **or** Conditional Access Administrator *(use with care; manage rights)*
- **Exchange Online**
  - View‑Only Organization Management
- **Microsoft Teams**
  - Teams Administrator *(or Communications Administrator for read scenarios)*
- **SharePoint / OneDrive**
  - SharePoint Administrator *(read tenant settings)*

> Notes:
> - Some Maester tests may require additional privileges depending on the tenant’s features.
> - If a test is *skipped/failed due to permissions*, grant the minimum read role for that service and re‑run.

---

## 6) Common Issues & Fixes

- **`Please run this script in an elevated session`** – Start PowerShell as *Administrator*.
- **`Install-Module` fails** – Ensure TLS 1.2 is enabled and that the PowerShell Gallery is reachable. Corporate proxies may require additional configuration.
- **Pester 3 vs 5 conflicts** – The installer ensures Pester 5+. If you still have issues, remove legacy Pester 3 from `$PSModulePath` or update it.
- **EXO/Teams/SPO not needed** – Disable the toggles at the top of `MaesterInteractive.ps1` to skip connecting to those services.

---

## 7) Parameters

### MaesterInteractive.ps1
- `-CustomerName <string>` – Overrides the default `ProtectionSets`.

### Install-MaesterPrereqs-AllUsers.ps1
- `-IncludeTeams` – Installs `MicrosoftTeams`.
- `-IncludeSPO` – Installs `PnP.PowerShell`.
- `-ForceReinstall` – Reinstalls modules even if the minimum version is present.
- `-VerboseLogging` – Enables verbose output.
- `-NoRolesPopup` – Suppresses the roles popup at the end.

---

## 8) Security Considerations

- Scripts install modules at **AllUsers** scope. Ensure you trust the source (PowerShell Gallery or your internal mirror).
- Prefer **read‑only** roles (e.g., Global Reader/Security Reader) for assessments.
- Avoid storing credentials; all connections are **interactive** by default.

---

## 9) Change Log (maintain as you iterate)

- **v1.1** – Added roles popup to prerequisites installer; README created; default customer set to `ProtectionSets`.
- **v1.0** – Initial scripts and Mapper‑aligned exports.

---

## 10) Support

- Re‑run the prerequisites installer with `-VerboseLogging` for diagnostics.
- Capture transcript logs if needed:
  ```powershell
  Start-Transcript -Path ./maester_setup.log -Force
  # run your commands
  Stop-Transcript
  ```

