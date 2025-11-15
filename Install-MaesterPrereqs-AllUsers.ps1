<#
.SYNOPSIS
Installs all prerequisites (for AllUsers) so the Maester→Protection Sets script runs smoothly,
then shows a popup with recommended admin roles for running Maester tests.

.PARAMETER IncludeTeams
Install MicrosoftTeams module as part of prerequisites. Default: $false.

.PARAMETER IncludeSPO
Install PnP.PowerShell (SharePoint) as part of prerequisites. Default: $false.

.PARAMETER ForceReinstall
Force reinstallation of target modules even if already present.

.PARAMETER VerboseLogging
Turns on detailed logging for troubleshooting.

.PARAMETER NoRolesPopup
Suppress the end-of-script popup that lists recommended admin roles.

.NOTES
- Run this script in an elevated (Administrator) PowerShell.
- Safe to re-run. Uses AllUsers scope so scheduled tasks and other users can run the main script.
#>

[CmdletBinding()]
param(
  [switch]$IncludeTeams = $false,
  [switch]$IncludeSPO   = $false,
  [switch]$ForceReinstall,
  [switch]$VerboseLogging,
  [switch]$NoRolesPopup
)

if ($VerboseLogging) { $PSDefaultParameterValues['*:Verbose'] = $true }

# -------------------- Helpers --------------------
function Write-Section([string]$Text) { Write-Host ("==> {0}" -f $Text) -ForegroundColor Cyan }
function Assert-Elevated {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) { throw "Please run this script in an elevated (Administrator) PowerShell session." }
}
function Ensure-Tls12 {
  try {
    if ([Net.ServicePointManager]::SecurityProtocol -band [Net.SecurityProtocolType]::Tls12) { return }
  } catch { }
  try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    Write-Verbose "Enabled TLS 1.2"
  } catch {
    Write-Warning "Could not enforce TLS 1.2: $($_.Exception.Message)"
  }
}
function Ensure-NuGetProvider {
  $nuget = Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue
  if (-not $nuget -or $nuget.Version -lt [Version]'2.8.5.201') {
    Write-Section "Installing/Updating NuGet provider"
    Install-PackageProvider -Name NuGet -Force -Scope AllUsers -ErrorAction Stop | Out-Null
  }
}
function Ensure-PSGallery {
  $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
  if (-not $repo) {
    Write-Section "Registering PowerShell Gallery"
    Register-PSRepository -Name PSGallery -SourceLocation 'https://www.powershellgallery.com/api/v2' `
      -InstallationPolicy Trusted -ErrorAction Stop
  } elseif ($repo.InstallationPolicy -ne 'Trusted') {
    Write-Section "Trusting PowerShell Gallery"
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
  }
}
function Install-ModuleSafe {
  param(
    [Parameter(Mandatory)][string]$Name,
    [string]$MinimumVersion,
    [switch]$ForceReinstall
  )
  $installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
  $need = $true
  if ($installed) {
    if ($MinimumVersion) {
      if ($installed.Version -ge [Version]$MinimumVersion -and -not $ForceReinstall) { $need = $false }
    } elseif (-not $ForceReinstall) { $need = $false }
  }
  if ($need) {
    Write-Section ("Installing module '{0}' (AllUsers){1}" -f $Name, $(if ($MinimumVersion) { " >= $MinimumVersion" } else { "" }))
    $params = @{
      Name          = $Name
      Scope         = 'AllUsers'
      Force         = $true
      AllowClobber  = $true
      ErrorAction   = 'Stop'
    }
    if ($MinimumVersion) { $params['MinimumVersion'] = $MinimumVersion }
    Install-Module @params
  } else {
    Write-Host ("- {0} already meets requirement{1}: {2}" -f $Name, $(if ($MinimumVersion) { " (>= $MinimumVersion)" } else { "" }), $installed.Version) `
      -ForegroundColor DarkGreen
  }
}
function Test-CommandAvailable {
  param([Parameter(Mandatory)][string[]]$Names)
  $allOk = $true
  foreach ($n in $Names) {
    if (-not (Get-Command $n -ErrorAction SilentlyContinue)) {
      Write-Warning "Command not found after install: $n"
      $allOk = $false
    } else {
      Write-Host ("✓ {0}" -f $n) -ForegroundColor Green
    }
  }
  return $allOk
}

function Show-RolesPopup {
  param(
    [string]$Title = 'Maester – Recommended Admin Roles',
    [string]$Customer = 'ProtectionSets'
  )
  try { Add-Type -AssemblyName System.Windows.Forms } catch {}

  $lines = @(
    "To run Maester tests with broad coverage (read/assessment), the following roles are recommended:",
    "",
    "Tenant-wide (Entra ID / M365):",
    " • Global Reader",
    " • Security Reader",
    " • Reports Reader",
    "Conditional Access / Identity (for CA policy reads):",
    " • Security Reader   (sufficient for read in many tenants)",
    " • OR Conditional Access Administrator (read/manage CA; use read-only where possible)",
    "Exchange Online:",
    " • View-Only Organization Management",
    "Microsoft Teams:",
    " • Teams Administrator (or Communications Administrator for read scenarios)",
    "SharePoint / OneDrive:",
    " • SharePoint Administrator (read-only access to tenant settings)",
    "",
    "Notes:",
    " • Some Maester tests may require higher privileges depending on tenant configuration and feature usage.",
    " • If a test is skipped/failed due to permissions, re-run after granting the specific service's read role.",
    "",
    "Customer context: " + $Customer
  ) -join [Environment]::NewLine

  [void][System.Windows.Forms.MessageBox]::Show(
    $lines,
    $Title,
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
  )
}

# -------------------- Execution --------------------
try {
  Assert-Elevated
  Ensure-Tls12
  Ensure-NuGetProvider
  Ensure-PSGallery

  # PowerShellGet v2 (optional but helps reliability on older hosts)
  try {
    $psget = Get-Module -ListAvailable -Name PowerShellGet | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $psget -or $psget.Version -lt [Version]'2.2.5') {
      Write-Section "Updating PowerShellGet to >= 2.2.5 (AllUsers)"
      Install-Module PowerShellGet -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
    }
  } catch {
    Write-Warning "Could not update PowerShellGet: $($_.Exception.Message)"
  }

  # --- Core modules ---
  Install-ModuleSafe -Name 'Pester'                   -MinimumVersion '5.5.0' -ForceReinstall:$ForceReinstall
  Install-ModuleSafe -Name 'Microsoft.Graph'          -MinimumVersion '2.0.0' -ForceReinstall:$ForceReinstall
  Install-ModuleSafe -Name 'ExchangeOnlineManagement' -MinimumVersion '3.0.0' -ForceReinstall:$ForceReinstall
  Install-ModuleSafe -Name 'Maester'                  -MinimumVersion '1.0.0' -ForceReinstall:$ForceReinstall

  # --- Optional modules ---
  if ($IncludeTeams) { Install-ModuleSafe -Name 'MicrosoftTeams' -MinimumVersion '5.0.0' -ForceReinstall:$ForceReinstall }
  if ($IncludeSPO)   { Install-ModuleSafe -Name 'PnP.PowerShell' -MinimumVersion '2.1.0' -ForceReinstall:$ForceReinstall }

  Write-Section "Validation: importing modules"
  Import-Module Pester -ErrorAction Stop
  Import-Module Microsoft.Graph -ErrorAction Stop
  Import-Module ExchangeOnlineManagement -ErrorAction Stop
  Import-Module Maester -ErrorAction Stop
  if ($IncludeTeams) { Import-Module MicrosoftTeams -ErrorAction Stop }
  if ($IncludeSPO)   { Import-Module PnP.PowerShell -ErrorAction Stop }

  Write-Section "Validation: commands available"
  $ok = $true
  $ok = Test-CommandAvailable -Names @(
    'Connect-MgGraph',
    'Invoke-Maester',
    'New-PesterConfiguration'
  ) -and $ok
  $ok = Test-CommandAvailable -Names @('Connect-ExchangeOnline') -and $ok
  if ($IncludeTeams) { $ok = Test-CommandAvailable -Names @('Connect-MicrosoftTeams') -and $ok }
  if ($IncludeSPO)   { $ok = Test-CommandAvailable -Names @('Connect-PnPOnline') -and $ok }

  if ($ok) {
    Write-Host "`nAll prerequisites installed and validated ✅" -ForegroundColor Green
    Write-Host "You can now run your main script (defaults to customer 'ProtectionSets')."
  } else {
    Write-Warning "Some commands were not found after installation. Check warnings above or re-run with -VerboseLogging."
  }

  if (-not $NoRolesPopup) {
    Show-RolesPopup -Customer 'ProtectionSets'
  }
}
catch {
  Write-Error "Failed to install prerequisites: $_"
  exit 1
}