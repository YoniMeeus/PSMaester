<#
.SYNOPSIS 
Interactive login Maester → Protection Sets run, aligned with DocuForgeMapper V3.
- Uses the customer selected in the mapper (no prompt) OR default below.
- Installs/uses Maester + dependencies (AllUsers scope).
- Produces detailed CSVs in <CustomerRoot>\ReportMapper so DocuForgeMapper V3 can map them.
- Also produces MaesterResults HTML/JSON alongside for reference.
Notes:
- Module installations use: -Scope AllUsers
- Script expects elevated PowerShell (Run as Administrator) to install modules for AllUsers.
#>
[CmdletBinding()]
param(
    # Default so the script runs unattended.
    [string]$CustomerName = 'ProtectionSets'
)

# ===================== ENABLE CONNECTIONS =====================
$EnableEXO   = $true
$EnableTeams = $false # optional
$EnableSPO   = $false # optional

# ===================== DEFAULTS =====================
$IncludeTags     = @()
$ExcludeTags     = @()
$DefaultSeverity = 'Informational' # guaranteed fallback so every row has a severity

# ===================== HELPERS =====================
function Write-Section([string]$Text) { Write-Host ("==> {0}" -f $Text) -ForegroundColor Cyan }
function Ensure-Folder([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path -Force -ErrorAction SilentlyContinue |
      Out-Null
  }
}
function Assert-Elevated {
  $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
  ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin) {
    throw "This operation requires an elevated PowerShell session (Run as Administrator) to install modules for AllUsers."
  }
}
function Ensure-ModulePresent {
  param([Parameter(Mandatory)][string[]]$Names)
  Assert-Elevated
  foreach ($n in $Names) {
    if (-not (Get-Module -ListAvailable -Name $n)) {
      Write-Section "Installing module '$n' (AllUsers)..."
      Install-Module -Name $n -Scope AllUsers -Force -ErrorAction Stop
    }
  }
}
function Ensure-MaesterModule {
  param([switch]$ForceFresh)
  Assert-Elevated
  $resolvedPath = $null
  if (-not $ForceFresh) {
    $mod = Get-Module Maester -ListAvailable |
      Sort-Object Version -Descending |
      Select-Object -First 1
    if ($mod) { $resolvedPath = $mod.ModuleBase }
  }
  if (-not $resolvedPath) {
    Write-Section "Installing Maester module (AllUsers)..."
    Install-Module -Name Maester -Scope AllUsers -Force -ErrorAction Stop
    $mod = Get-Module Maester -ListAvailable |
      Sort-Object Version -Descending |
      Select-Object -First 1
    if (-not $mod) { throw "Failed to install Maester for AllUsers." }
    $resolvedPath = $mod.ModuleBase
  }
  $psd1 = Join-Path $resolvedPath 'Maester.psd1'
  if (Test-Path -LiteralPath $psd1) { Import-Module $psd1 -Force -ErrorAction Stop }
  else { Import-Module $resolvedPath -Force -ErrorAction Stop }
  [pscustomobject]@{ Path = $resolvedPath; WasDownloaded = $false }
}
function Ensure-GraphModules {
  if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
    Ensure-ModulePresent -Names 'Microsoft.Graph'
  }
}
function Install-MaesterTestsTemp {
  param([Parameter(Mandatory)][string]$TargetRoot)
  if (Test-Path -LiteralPath $TargetRoot) {
    Remove-Item -LiteralPath $TargetRoot -Recurse -Force -ErrorAction SilentlyContinue
  }
  Ensure-Folder $TargetRoot
  Push-Location $TargetRoot
  try { Install-MaesterTests } finally { Pop-Location }
  return $TargetRoot
}

# ===================== AST / Tag Helpers =====================
function Get-ItBlocksMetaFromFile {
  [CmdletBinding()] param([Parameter(Mandatory)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Test file not found: $Path" }
  $tokens=$null; $errors=$null
  $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
  $itCmds = $ast.FindAll({
    param($node)
    if ($node -is [System.Management.Automation.Language.CommandAst]) {
      try { $node.GetCommandName() -eq 'It' } catch { $false }
    } else { $false }
  }, $true)
  foreach ($cmd in $itCmds) {
    $title=$null
    foreach ($elem in $cmd.CommandElements) {
      if ($elem -is [System.Management.Automation.Language.CommandParameterAst]) { continue }
      if ($elem -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $title = $elem.Value; break }
    }
    $tags = [System.Collections.Generic.List[string]]::new()
    for ($i=0; $i -lt $cmd.CommandElements.Count; $i++) {
      $elem = $cmd.CommandElements[$i]
      if ($elem -is [System.Management.Automation.Language.CommandParameterAst] -and $elem.ParameterName -eq 'Tag') {
        if ($i + 1 -lt $cmd.CommandElements.Count) {
          $arg = $cmd.CommandElements[$i + 1]
          switch ($arg.GetType().FullName) {
            'System.Management.Automation.Language.StringConstantExpressionAst' { $tags.Add($arg.Value) > $null }
            'System.Management.Automation.Language.ExpandableStringExpressionAst' { $tags.Add($arg.Value) > $null }
            'System.Management.Automation.Language.ArrayLiteralAst' {
              foreach ($el in $arg.Elements) {
                if ($el -is [System.Management.Automation.Language.StringConstantExpressionAst]) { $tags.Add($el.Value) > $null }
                elseif ($el -is [System.Management.Automation.Language.ExpandableStringExpressionAst]) { $tags.Add($el.Value) > $null }
              }
            }
          }
        }
      }
    }
    $sbAst = $cmd.CommandElements `
      | Where-Object { $_ -is [System.Management.Automation.Language.ScriptBlockExpressionAst] } `
      | Select-Object -Last 1 -ExpandProperty ScriptBlock
    $fnCalls = [System.Collections.Generic.List[string]]::new()
    if ($sbAst) {
      try {
        $callCmds = $sbAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)
        foreach ($c in $callCmds) {
          $name = $c.GetCommandName()
          if ($name -and ($name -match '^(Test-ORCA|Test-MtCisa|Test-MtEidsca|Test-MtCis|Test-Mt)')) {
            $fnCalls.Add($name) > $null
          }
        }
      } catch {}
    }
    $docUrl = $null
    if ($title) { $m = [regex]::Match($title, 'https?://\S+'); if ($m.Success) { $docUrl = $m.Value } }
    $titleClean = if ($title) { ($title -replace '\s+See\s+https?://\S+','').Trim() } else { $null }
    [pscustomobject]@{
      Title     = $titleClean
      Tags      = @($tags)
      Functions = @($fnCalls)
      DocUrl    = $docUrl
      SourceFile= $Path
    }
  }
}
function Build-ItIndex { param([Parameter(Mandatory)][string]$TestsRoot)
  $files = Get-ChildItem -Path (Join-Path $TestsRoot '*') -Recurse -Include *.Tests.ps1 -File -ErrorAction SilentlyContinue
  $index = [System.Collections.Generic.List[object]]::new()
  foreach ($f in $files) { foreach ($r in (Get-ItBlocksMetaFromFile -Path $f.FullName)) { $index.Add($r) } }
  return $index
}
# -- Code patterns to match --
$rxParts = @(
  'MT\.\d{3,5}(?:\.[A-Za-z0-9]+)*',
  'EIDSCA\.[A-Z]{2}\d{2}',
  'CISA\.MS\.[A-Z]+\.\d+(?:\.\d+)?',
  'CIS\.M365\.\d+\.\d+(?:\.\d+)?',
  'ORCA\.\d+(?:\.\d+)?',
  'Test-ORC[A-Za-z0-9._-]+',
  'Test-Mt[A-Za-z0-9._-]+'
)
# FIX: proper alternation (no literal newlines)
$rxCode = '(?im)(' + ($rxParts -join '|') + ')'

# ===================== Severity helpers =====================
function Normalize-Severity { param([string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
  $v = $Value.Trim()
  switch -Regex ($v) {
    '^(crit(ical)?)$'       { return 'Critical' }
    '^(high)$'              { return 'High' }
    '^(med(ium)?)$'         { return 'Medium' }
    '^(mod(erate)?)$'       { return 'Moderate' }
    '^(low)$'               { return 'Low' }
    '^(info|informational)$'{ return 'Informational' }
    default                 { return ($v.Substring(0,1).ToUpper() + $v.Substring(1).ToLower()) }
  }
}
function Get-SeverityFromTagsOrText {
  param([string[]]$Tags,[string]$Title,[string]$Context)
  $normTags = @(); foreach ($t in ($Tags ?? @())) { if ($t) { $s = [string]$t; $s = $s.Trim(); if ($s) { $normTags += $s } } }
  $hay = ($normTags -join '; ')
  if ($Title)   { $hay = "$hay; $Title" }
  if ($Context) { $hay = "$hay; $Context" }
  $hayL = $hay.ToLowerInvariant()
  $m = [regex]::Match($hayL, '\b(severity|sev)\s*[:=\-_\ ]\s*(critical|high|medium|moderate|low|info|informational)\b')
  if ($m.Success) { return Normalize-Severity $m.Groups[2].Value }
  foreach ($lvl in 'critical','high','medium','moderate','low','info','informational') {
    if ($hayL -match "(^|[;,\s\(\)\[\]\-_/])$lvl($|[;,\s\(\)\[\]\-_/])") {
      return Normalize-Severity $lvl
    }
  }
  return $null
}
function Build-SeverityMapFromJson { param([Parameter(Mandatory)][string]$JsonPath)
  $map = @{}
  if (-not (Test-Path -LiteralPath $JsonPath -PathType Leaf)) { return $map }
  try { $json = Get-Content -LiteralPath $JsonPath -Raw | ConvertFrom-Json -Depth 100 } catch { return $map }
  function Add-Entry([hashtable]$m,[string]$name,[string]$path,[string]$sev) {
    if ([string]::IsNullOrWhiteSpace($sev)) { return }
    $val = Normalize-Severity $sev
    if (-not $val) { return }
    if ($name) { $m[$name.ToLowerInvariant()] = $val }
    if ($name -and $path) { $m[("$path::" + $name).ToLowerInvariant()] = $val }
  }
  function Visit($node,[string]$currentPath) {
    if ($null -eq $node) { return }
    if ($node -is [System.Collections.IEnumerable] -and -not ($node -is [string])) {
      foreach ($it in $node) { Visit -node $it -currentPath $currentPath }
      return
    }
    $pso = [pscustomobject]$node
    $props = $pso.PSObject.Properties
    $name = $null; $path = $currentPath; $sev = $null
    foreach ($pn in @('Name','Title','TestName','DisplayName')) { if ($props[$pn]) { $n=[string]$props[$pn].Value; if ($n){ $name=$n; break } } }
    foreach ($sn in @('Severity','severity','Risk','risk','Level','level')) { if ($props[$sn]) { $s=[string]$props[$sn].Value; if ($s){ $sev=$s; break } } }
    foreach ($ppn in @('Path','ScriptPath','File','Source','SourceFile')) { if (-not $path -and $props[$ppn]) { $path=[string]$props[$ppn].Value } }
    if ($name -and $sev) { Add-Entry -m $map -name $name -path $path -sev $sev }
    foreach ($cn in @('TestResult','Tests','Items','Children','Results','value')) {
      if ($props[$cn]) { Visit -node $props[$cn].Value -currentPath $path }
    }
    foreach ($pp in $props) {
      $val = $pp.Value
      if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) { Visit -node $val -currentPath $path }
      elseif ($val -is [psobject]) { Visit -node $val -currentPath $path }
    }
  }
  Visit -node $json -currentPath $null
  return $map
}

# ===================== RESOLVE CUSTOMER (default already set) =====================
if ([string]::IsNullOrWhiteSpace($CustomerName)) {
  # Absolute fallback (shouldn’t happen as we set a default)
  $CustomerName = 'ProtectionSets'
}
$env:DSC_SelectedCustomer = $CustomerName  # keep env in sync for downstream tools

# ===================== PATHS (aligned with Mapper V3) =====================
# <CustomerRoot> = C:\M365Factory\Customers\<CustomerName>
$CustomerRoot    = Join-Path 'C:\M365Factory\Customers' $CustomerName
$ReportMapperDir = Join-Path $CustomerRoot 'ReportMapper'
# We place Maester artifacts here, because Mapper V3’s "ReportMapper" button expects the inputs in this folder.
# (maester_failed_detailed.csv, maester_skipped_detailed.csv)
# It will then generate MaesterTagMatches.html/.csv in the same folder.
$runRoot   = $CustomerRoot
$testsRoot = Join-Path $runRoot 'maester-tests'
$resultsRoot = $ReportMapperDir
Ensure-Folder $CustomerRoot
Ensure-Folder $resultsRoot

Write-Host ""
Write-Host ("========== Maester→PS Matcher Run for Customer '{0}' ==========" -f $CustomerName) -ForegroundColor DarkCyan
try {
  Write-Host ""
  $runId = Get-Date -Format 'yyyyMMdd_HHmmss'
  Write-Host ("========== Maester→PS Matcher Run {0} ==========" -f $runId) -ForegroundColor DarkCyan

  Write-Section 'Ensuring required modules'
  Ensure-GraphModules
  $mods = @()
  if ($EnableEXO)   { $mods += 'ExchangeOnlineManagement' }
  if ($EnableTeams) { $mods += 'MicrosoftTeams' }
  if ($EnableSPO)   { $mods += 'PnP.PowerShell' }
  if ($mods.Count)  { Ensure-ModulePresent -Names $mods }

  Write-Section 'Ensuring Maester module'
  $mm = Ensure-MaesterModule
  Write-Host ("Maester module path: {0}" -f $mm.Path) -ForegroundColor DarkGreen

  Write-Section 'Installing Maester tests to run folder'
  $testsRoot = Install-MaesterTestsTemp -TargetRoot $testsRoot

  # ---------- INTERACTIVE AUTH ----------
  Write-Section 'Connecting to cloud services (interactive login)'
  # Graph
  try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
  Connect-MgGraph -Scopes `
    "User.Read.All",
    "Directory.Read.All",
    "DeviceManagementConfiguration.Read.All",
    "DeviceManagementManagedDevices.Read.All",
    "DirectoryRecommendations.Read.All",
    "IdentityRiskEvent.Read.All",
    "Policy.Read.ConditionalAccess",
    "PrivilegedAccess.Read.AzureAD",
    "Reports.Read.All",
    "RoleEligibilitySchedule.Read.Directory",
    "RoleManagement.Read.All",
    "SharePointTenantSettings.Read.All"
  Write-Host "Connected: Microsoft Graph (interactive)" -ForegroundColor DarkGreen

  # Exchange Online
  if ($EnableEXO) {
    try { Connect-ExchangeOnline -ShowBanner:$false; Write-Host "Connected: Exchange Online (interactive)" -ForegroundColor DarkGreen }
    catch { Write-Warning "EXO interactive connection failed: $($_.Exception.Message)" }
  }

  # Teams
  if ($EnableTeams) {
    try { Import-Module MicrosoftTeams -Force; Connect-MicrosoftTeams; Write-Host "Connected: Microsoft Teams (interactive)" -ForegroundColor DarkGreen }
    catch { Write-Warning "Teams interactive connection failed: $($_.Exception.Message)" }
  }

  # ---------- RUN MAESTER ----------
  Write-Section 'Running Maester tests'
  $outHtml = Join-Path $resultsRoot 'MaesterResults.html' # reference output
  $outJson = Join-Path $resultsRoot 'MaesterResults.json' # used to enrich severities
  $cfg = New-PesterConfiguration
  $cfg.Run.Path = $testsRoot
  if ($IncludeTags) { $cfg.Filter.Tag        = $IncludeTags }
  if ($ExcludeTags) { $cfg.Filter.ExcludeTag = $ExcludeTags }

  # Keep your original switches and CSV export; we place everything into $resultsRoot.
  $run = Invoke-Maester -PesterConfiguration $cfg `
    -OutputHtmlFile $outHtml `
    -OutputJsonFile $outJson `
    -OutputFolder   $resultsRoot `
    -PassThru -NonInteractive -NoLogo -Verbosity Normal -ExportCsv

  Write-Section 'Indexing test files (AST) for Title/Tags/Alias/DocUrl'
  $itIndex = Build-ItIndex -TestsRoot $testsRoot
  $itByFileTitle = @{}
  foreach ($r in $itIndex) {
    $keyTitle = if ($r.Title) { $r.Title } else { '' }
    $key = ($r.SourceFile + '::' + $keyTitle).ToLowerInvariant()
    if ($key) { $itByFileTitle[$key] = $r }
  }
  function Get-CodesFromMeta { param([object]$ItMeta)
    if (-not $ItMeta) { return @() }
    $c=@()
    foreach ($t in @($ItMeta.Tags))     { $s=($t -as [string]).Trim(); if ($s -match $rxCode) { $c += $s.ToUpper() } }
    foreach ($fn in @($ItMeta.Functions)){ $s=($fn -as [string]).Trim(); if ($s -match $rxCode) { $c += $s.ToUpper() } }
    $c | Sort-Object -Unique
  }

  # Build Severity map from Maester JSON
  Write-Section 'Building Severity map from JSON'
  $severityMap = Build-SeverityMapFromJson -JsonPath $outJson

  $failedRows  = [System.Collections.Generic.List[object]]::new()
  $skippedRows = [System.Collections.Generic.List[object]]::new()
  $testItems = if ($run -and $run.TestResult) { $run.TestResult }
               elseif ($run -and $run.Tests)   { $run.Tests }
               else { @() }

  foreach ($t in $testItems) {
    $status   = $t.Result; if (-not $status -and ($t.Status)) { $status = $t.Status }
    $name     = $t.Name
    $err      = $t.ErrorRecord
    $stack    = $t.StackTrace
    $duration = $t.Time
    $context  = $t.Context
    $source   = $t.Path
    if (-not $source -and $t.ScriptBlock -and $t.ScriptBlock.File) { $source = $t.ScriptBlock.File }

    $nm   = if ($name) { $name } else { '' }
    $meta = $null
    if ($source) {
      $metaKey = ($source + '::' + $nm).ToLowerInvariant()
      if ($itByFileTitle.ContainsKey($metaKey)) { $meta = $itByFileTitle[$metaKey] }
    }
    if (-not $meta) { $meta = $itIndex | Where-Object { $_.Title -eq $name } | Select-Object -First 1 }

    $codesMeta  = Get-CodesFromMeta -ItMeta $meta
    $codesTitle = if ($name) { [regex]::Matches($name, $rxCode) | ForEach-Object { $_.Value.ToUpperInvariant() } } else { @() }
    $codesUnion = @($codesMeta + $codesTitle) | Sort-Object -Unique
    $doc  = if ($meta) { $meta.DocUrl } else { $null }
    $tags = if ($meta) { ($meta.Tags -join '; ') } else { $null }
    $errMsg = if ($err) { ($err.Exception.Message -replace '\r?\n',' ') } else { $null }

    # ---- Severity resolution chain ----
    $severity = $null
    if ($t.PSObject.Properties['Severity'] -and $t.Severity) { $severity = Normalize-Severity $t.Severity }
    if (-not $severity) {
      $k1 = if ($source) { ($source + '::' + $nm).ToLowerInvariant() } else { $null }
      $k2 = if ($nm) { $nm.ToLowerInvariant() } else { $null }
      if ($k1 -and $severityMap.ContainsKey($k1)) { $severity = $severityMap[$k1] }
      if (-not $severity -and $k2 -and $severityMap.ContainsKey($k2)) { $severity = $severityMap[$k2] }
    }
    if (-not $severity) {
      $tagsArray = @(); if ($meta -and $meta.Tags) { $tagsArray = @($meta.Tags) }
      $severity = Get-SeverityFromTagsOrText -Tags $tagsArray -Title $name -Context $context
    }
    if (-not $severity) { $severity = $DefaultSeverity }
    # ----------------------------------

    $row = [pscustomobject]@{
      Status     = $status
      Title      = $name
      Tags       = $tags
      Codes      = ($codesUnion -join '; ')
      DocUrl     = $doc
      SourceFile = $source
      Error      = $errMsg
      Stack      = $stack
      Duration   = $duration
      Context    = $context
      Severity   = $severity
    }
    if     ($status -eq 'Failed')  { $failedRows.Add($row)  > $null }
    elseif ($status -eq 'Skipped') { $skippedRows.Add($row) > $null }
  }

  # Export CSVs to the exact folder/filenames the Mapper wants
  $failedCsv  = Join-Path $resultsRoot 'maester_failed_detailed.csv'
  $skippedCsv = Join-Path $resultsRoot 'maester_skipped_detailed.csv'
  $failedRows  | Export-Csv -Path $failedCsv  -NoTypeInformation -Encoding UTF8
  $skippedRows | Export-Csv -Path $skippedCsv -NoTypeInformation -Encoding UTF8

  Write-Host ""
  Write-Host "Saved detailed CSVs for Mapper:" -ForegroundColor Green
  Write-Host " - $failedCsv"
  Write-Host " - $skippedCsv"
  Write-Host ""
  Write-Host "Reference outputs:" -ForegroundColor Green
  Write-Host " - $outHtml"
  Write-Host " - $outJson"
}
catch {
  Write-Error "An error occurred during the Maester→PS Matcher run: $_"
}
finally {
  Write-Section ("Run completed at {0}" -f (Get-Date))
}

Disconnect-Graph
Disconnect-MgGraph

# -- Added: OK dialog then hard exit (no logic changes to preceding code) --
try { Add-Type -AssemblyName System.Windows.Forms } catch {}
try {
  [void][System.Windows.Forms.MessageBox]::Show(
    "Script finished. Click OK to close.",
    "<ScriptTitle>", # MeasterInteractive or MeasterSSL
    'OK',
    'Information'
  )
} catch {}
$host.SetShouldExit(0)
[Environment]::Exit(0)
