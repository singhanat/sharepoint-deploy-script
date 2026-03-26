# =========================
# export.ps1 (CONFIG + PnP Provisioning Template + TitleRequired map + selective data export)
# =========================
# Run: .\export.ps1
# Reads: config.json (same folder)
# Outputs: <ScriptDir>\<SiteNameFromUrl>\
#   - template.xml
#   - lists_index.csv
#   - document_libraries.txt
#   - title_required.csv
#   - data_<ListTitle>.csv  (only lists in export.exportDataLists)
#   - data_index.csv
#   - export_run.log
# =========================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-PSMajor { try { [int]$PSVersionTable.PSVersion.Major } catch { 5 } }
function Get-ScriptDir { if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path } }
function Sanitize-FileName {
  param([string]$Name)
  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($ch in $invalid) { $Name = $Name.Replace($ch, "_") }
  $Name = ($Name -replace '\s+', ' ').Trim()
  if ([string]::IsNullOrWhiteSpace($Name)) { $Name = "unnamed" }
  $Name
}
function Get-SiteFolderNameFromUrl {
  param([string]$Url)
  try {
    $u = [System.Uri]$Url
    $segments = @($u.AbsolutePath.Trim("/") -split "/") | Where-Object { $_ }
    if ($segments.Count -gt 0) { return (Sanitize-FileName $segments[-1]) }
  }
  catch {}
  return (Sanitize-FileName (($Url.TrimEnd("/") -split "/")[-1]))
}
function Read-Config {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "Config file not found: $Path" }
  $raw = Get-Content -Raw -Path $Path
  return ($raw | ConvertFrom-Json)
}
function Get-PropSafe {
  param($Obj, [string]$Name, $Default = $null)
  if ($null -eq $Obj) { return $Default }
  $p = $Obj.PSObject.Properties[$Name]
  if ($null -eq $p) { return $Default }
  $p.Value
}

# --- read config ---
$scriptDir = Get-ScriptDir
$configPath = Join-Path $scriptDir "config.json"
$cfg = Read-Config $configPath

$SiteUrl = [string]$cfg.export.siteUrl
if ([string]::IsNullOrWhiteSpace($SiteUrl)) { throw "config.json: export.siteUrl is required" }

$IncludeHiddenLists = [bool]$cfg.export.includeHiddenLists
$IncludeSystemLists = [bool]$cfg.export.includeSystemLists

$ExcludeSiteUsers = $true
if ($null -ne $cfg.export.excludeSiteUsers) {
  $ExcludeSiteUsers = [bool]$cfg.export.excludeSiteUsers
}

$ExportLists = @()
try { $ExportLists = @( $cfg.export.exportLists | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } ) } catch { $ExportLists = @() }

$ExportDataLists = @()
try { $ExportDataLists = @($cfg.export.exportDataLists) } catch { $ExportDataLists = @() }
$DataPageSize = 2000
try {
  if ($cfg.export.dataPageSize) { $DataPageSize = [int]$cfg.export.dataPageSize }
}
catch {}

# --- output/log ---
$siteFolder = Get-SiteFolderNameFromUrl $SiteUrl
$outDir = Join-Path $scriptDir $siteFolder
if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$logPath = Join-Path $outDir "export_run.log"
if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force -ErrorAction SilentlyContinue }
function Log {
  param([string]$Msg, [string]$Level = "INFO")
  $line = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff"), $Level, $Msg
  Add-Content -Path $logPath -Value $line -Encoding UTF8
  Write-Host $line
}

# --- PnP abstraction ---
$script:UseLegacyPnP = $false
function Ensure-PnPModule {
  $major = Get-PSMajor
  if ($major -ge 7) {
    $script:UseLegacyPnP = $false
    if (-not (Get-Module -ListAvailable -Name "PnP.PowerShell" | Select-Object -First 1)) {
      Log "PnP.PowerShell not found. Installing for current user..." "WARN"
      Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module PnP.PowerShell -ErrorAction Stop
    Log "Using module: PnP.PowerShell (PS $major)" "INFO"
  }
  else {
    $script:UseLegacyPnP = $true
    $legacyName = "SharePointPnPPowerShellOnline"
    if (-not (Get-Module -ListAvailable -Name $legacyName | Select-Object -First 1)) {
      Log "$legacyName not found. Installing for current user..." "WARN"
      Install-Module $legacyName -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $legacyName -ErrorAction Stop
    Log "Using module: $legacyName (Windows PowerShell 5.1)" "INFO"
  }
}
function PnP-Connect([string]$Url) {
  if ($script:UseLegacyPnP) {
    Log "Connect-PnPOnline -UseWebLogin => $Url" "INFO"
    Connect-PnPOnline -Url $Url -UseWebLogin -WarningAction SilentlyContinue
  }
  else {
    Log "Connect-PnPOnline -Interactive => $Url" "INFO"
    Connect-PnPOnline -Url $Url -Interactive
  }
}
function PnP-Disconnect { try { Disconnect-PnPOnline | Out-Null } catch {} }

function Get-ListsFiltered {
  param([bool]$IncludeHidden, [bool]$IncludeSystem)
  $lists = Get-PnPList -ErrorAction Stop | Sort-Object Id -Unique

  if (-not $IncludeHidden) { $lists = $lists | Where-Object { $_.Hidden -ne $true } }

  if (-not $IncludeSystem) {
    $systemTitles = @(
      "Style Library", "Form Templates", "Preservation Hold Library", "Site Assets", "Site Pages",
      "Pages", "MicroFeed", "Master Page Gallery", "Theme Gallery", "Composed Looks",
      "Site Collection Documents", "Site Collection Images", "Workflow Tasks", "Workflow History"
    )
    $lists = $lists | Where-Object { $systemTitles -notcontains $_.Title }
  }
  return $lists
}

function Export-ProvisioningTemplate {
  param([string]$OutPath, [string[]]$ListsToExtract)
  # ปรับ Handler ให้โฟกัสแค่ Lists กับ Fields (Views จะมาพร้อมกับ Lists อัตโนมัติ)
  # เหมือนแนวคิดของน้อง เพื่อลดความซับซ้อนและหลีกเลี่ยง Access Denied
  $handlers = "Lists,Fields"
  if ($script:UseLegacyPnP) {
    Log "Exporting template (legacy) -> $OutPath" "INFO"
    Get-PnPProvisioningTemplate -Out $OutPath -Handlers $handlers -ListsToExtract $ListsToExtract -ErrorAction Stop
  }
  else {
    Log "Exporting template (modern) -> $OutPath" "INFO"
    Get-PnPSiteTemplate -Out $OutPath -Handlers $handlers -ListsToExtract $ListsToExtract -ErrorAction Stop
  }
}

function Export-TitleRequiredMap {
  param($Lists, [string]$OutPath)

  $rows = New-Object System.Collections.Generic.List[object]
  foreach ($lst in $Lists) {
    if ([int]$lst.BaseTemplate -ne 100) { continue }
    try {
      $f = Get-PnPField -List $lst -Identity "Title" -ErrorAction Stop
      $req = $false
      try { $req = [bool]$f.Required } catch { $req = $false }
      $rows.Add([pscustomobject]@{ ListTitle = $lst.Title; TitleRequired = $req })
    }
    catch {}
  }

  $rows | Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8
  Log ("Wrote: {0} (count={1})" -f $OutPath, $rows.Count) "INFO"
}

function Get-ExportableDataFields {
  param($List)

  $fields = Get-PnPField -List $List -ErrorAction Stop
  $keep = New-Object System.Collections.Generic.List[string]

  $skipInternal = @(
    "ContentType", "ContentTypeId", "Attachments", "FileLeafRef", "FileRef", "FileDirRef",
    "Modified", "Created", "Author", "Editor", "GUID", "_UIVersionString", "_UIVersion",
    "AppAuthor", "AppEditor", "Edit", "LinkTitle", "LinkTitleNoMenu", "LinkFilename", "LinkFilenameNoMenu",
    "DocIcon", "ComplianceAssetId", "ID" # ID not needed for import (we will create new items)
  )

  foreach ($f in $fields) {
    $hidden = [bool](Get-PropSafe $f "Hidden" $false)
    $readonly = [bool](Get-PropSafe $f "ReadOnlyField" $false)
    $sealed = [bool](Get-PropSafe $f "Sealed" $false)
    $internal = [string](Get-PropSafe $f "InternalName" "")
    $typeStr = [string](Get-PropSafe $f "TypeAsString" "")

    if ($hidden -or $readonly -or $sealed) { continue }
    if ([string]::IsNullOrWhiteSpace($internal)) { continue }
    if ($skipInternal -contains $internal) { continue }

    # keep common, importable types
    if ($typeStr -match 'Text|Note|Number|Currency|DateTime|Boolean|Choice|User|Lookup|URL|Guid') {
      $keep.Add($internal)
    }
  }

  # Include Title if present and not excluded
  if ($keep -notcontains "Title") {
    # If the list has Title field and it's not sealed/hidden, the loop likely added it already.
    # But some lists may have internal name "Title" still add.
    $keep.Add("Title")
  }

  return $keep | Select-Object -Unique
}

function Export-ListDataToCsv {
  param(
    [Parameter(Mandatory = $true)][string]$ListTitle,
    [Parameter(Mandatory = $true)][string]$OutPath,
    [int]$PageSize = 2000
  )

  $lst = Get-PnPList -Identity $ListTitle -ErrorAction Stop

  if ([int]$lst.BaseTemplate -eq 101) {
    Log "Skip data export for Document Library: $ListTitle" "WARN"
    return @{ Rows = 0; File = $OutPath; Skipped = $true }
  }

  $fields = Get-ExportableDataFields -List $lst
  Log ("Data export fields [{0}] = {1}" -f $ListTitle, ($fields -join ", ")) "INFO"

  $rows = New-Object System.Collections.Generic.List[object]
  $items = Get-PnPListItem -List $lst -PageSize $PageSize -Fields $fields -ErrorAction Stop

  foreach ($it in $items) {
    $o = [ordered]@{}
    foreach ($fn in $fields) {
      $val = $null
      if ($it.FieldValues -and $it.FieldValues.ContainsKey($fn)) { $val = $it.FieldValues[$fn] }

      # flatten (best-effort)
      if ($null -eq $val) {
        $o[$fn] = $null
      }
      elseif ($val -is [System.Array]) {
        $o[$fn] = (($val | ForEach-Object { "$_" }) -join "|")
      }
      else {
        $tn = $val.GetType().FullName
        if ($tn -like "*FieldUserValue") {
          $email = $null; $lookup = $null
          try { $email = $val.Email } catch {}
          try { $lookup = $val.LookupValue } catch {}
          $o[$fn] = $(if ($email) { $email } else { $lookup })
        }
        elseif ($tn -like "*FieldLookupValue") {
          try { $o[$fn] = $val.LookupValue } catch { $o[$fn] = "$val" }
        }
        elseif ($tn -like "*FieldUrlValue") {
          try { $o[$fn] = $val.Url } catch { $o[$fn] = "$val" }
        }
        else {
          $o[$fn] = $val
        }
      }
    }
    $rows.Add([pscustomobject]$o)
  }

  $rows | Export-Csv -Path $OutPath -NoTypeInformation -Encoding UTF8
  return @{ Rows = $rows.Count; File = $OutPath; Skipped = $false }
}

# -------------------- MAIN --------------------
Log "Config: $configPath" "INFO"
Log "Output folder: $outDir" "INFO"
Log ("ExportDataLists: {0}" -f (($ExportDataLists | ForEach-Object { "$_" }) -join ", ")) "INFO"

Ensure-PnPModule
PnP-Connect $SiteUrl

$lists = @(Get-ListsFiltered -IncludeHidden:$IncludeHiddenLists -IncludeSystem:$IncludeSystemLists)

if ($ExportLists.Count -gt 0) {
  $lists = @( $lists | Where-Object { $ExportLists -contains $_.Title } )
}

Log ("Lists selected: {0}" -f $lists.Count) "INFO"

$indexPath = Join-Path $outDir "lists_index.csv"
($lists | ForEach-Object {
  [pscustomobject]@{
    Title        = $_.Title
    Id           = $_.Id
    BaseTemplate = $_.BaseTemplate
    Hidden       = $_.Hidden
  }
}) | Export-Csv -Path $indexPath -NoTypeInformation -Encoding UTF8
Log "Wrote: $indexPath" "INFO"

$docLibPath = Join-Path $outDir "document_libraries.txt"
$docLibTitles = [string[]]@( $lists | Where-Object { [int]$_.BaseTemplate -eq 101 } | Select-Object -ExpandProperty Title | Where-Object { $_ } | Sort-Object -Unique )
[System.IO.File]::WriteAllLines($docLibPath, $docLibTitles, (New-Object System.Text.UTF8Encoding($true)))
Log ("Wrote: {0} (count={1})" -f $docLibPath, $docLibTitles.Count) "INFO"

$templatePath = Join-Path $outDir "template.xml"
$listsToExtract = [string[]]@( $lists | Select-Object -ExpandProperty Title | Where-Object { $_ } | Sort-Object -Unique )
Export-ProvisioningTemplate -OutPath $templatePath -ListsToExtract $listsToExtract
Log "Wrote: $templatePath" "INFO"

if ($ExcludeSiteUsers -or $true) {
  try {
    $xmlContent = Get-Content -Raw -Path $templatePath -Encoding UTF8
    $xml = [xml]$xmlContent
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $nsmgr.AddNamespace("pnp", $xml.DocumentElement.NamespaceURI)
    $saveXml = $false

    if ($ExcludeSiteUsers) {
      $nodesToRemove = $xml.SelectNodes("//pnp:Security/pnp:AdditionalAdministrators | //pnp:Security/pnp:AdditionalOwners | //pnp:Security/pnp:AdditionalMembers | //pnp:SiteGroups/pnp:SiteGroup/pnp:Members", $nsmgr)
      $removedCount = 0
      foreach ($node in $nodesToRemove) {
        $node.ParentNode.RemoveChild($node) | Out-Null
        $removedCount++
      }
      if ($removedCount -gt 0) {
        $saveXml = $true
        Log ("Removed {0} user/member nodes from template.xml" -f $removedCount) "INFO"
      }
    }

    # Clean up problematic nodes that usually cause Access is denied at import
    $elementsToRemove = @("Webhooks", "PropertyBagEntries", "Security")
    $cleanedCount = 0
    foreach ($elemName in $elementsToRemove) {
      $nodesToClean = $xml.SelectNodes("//*[local-name()='$elemName']")
      foreach ($node in $nodesToClean) {
        if ($node -and $node.ParentNode) {
          $node.ParentNode.RemoveChild($node) | Out-Null
          $cleanedCount++
        }
      }
    }
    if ($cleanedCount -gt 0) {
      Log ("Cleaned up {0} problematic tags (Webhooks, Security, PropertyBag) from template." -f $cleanedCount) "INFO"
      $saveXml = $true
    }

    # Clean up system Fields generally (SiteFields & List Fields)
    $systemFieldIds = New-Object System.Collections.Generic.List[string]
    $systemFieldNames = New-Object System.Collections.Generic.List[string]
    $allFields = $xml.SelectNodes("//*[local-name()='Field']")
    $removedFields = 0
    foreach ($node in $allFields) {
      if ($node -and $node.ParentNode) {
        $sourceId = $node.GetAttribute("SourceID")
        $name = $node.GetAttribute("Name")
        $group = $node.GetAttribute("Group")
        $id = $node.GetAttribute("ID")
        
        # Identify system fields
        if ($sourceId -eq "http://schemas.microsoft.com/sharepoint/v3" -or 
          $name -match "^_|^TaxCatchAll|^Media|^ComplianceTag|^TriggerFlowInfo|^A2ODMountCount" -or 
          $group -eq "_Hidden" -or 
          $group -match "Document and Record Management") {
            
          if (-not [string]::IsNullOrWhiteSpace($id)) { $systemFieldIds.Add($id) }
          if (-not [string]::IsNullOrWhiteSpace($name)) { $systemFieldNames.Add($name) }
          
          $node.ParentNode.RemoveChild($node) | Out-Null
          $removedFields++
          $saveXml = $true
        }
      }
    }
    if ($removedFields -gt 0) {
      Log ("Removed {0} system Fields from template." -f $removedFields) "INFO"
      
      # Clean up FieldRefs that point to removed fields
      $removedFieldRefs = 0
      $allFieldRefs = $xml.SelectNodes("//*[local-name()='FieldRef']")
      foreach ($ref in $allFieldRefs) {
        if ($ref -and $ref.ParentNode) {
          $refName = $ref.GetAttribute("Name")
          $refId = $ref.GetAttribute("ID")
          
          if ($systemFieldNames.Contains($refName) -or $systemFieldIds.Contains($refId)) {
            $ref.ParentNode.RemoveChild($ref) | Out-Null
            $removedFieldRefs++
            $saveXml = $true
          }
        }
      }
      if ($removedFieldRefs -gt 0) {
        Log ("Removed {0} system FieldRefs from template." -f $removedFieldRefs) "INFO"
      }
    }

    if ($saveXml) {
      $xml.Save($templatePath)
    }
  }
  catch {
    Log "Failed to strip elements from template: $($_.Exception.Message)" "WARN"
  }
}

$titleReqPath = Join-Path $outDir "title_required.csv"
Export-TitleRequiredMap -Lists $lists -OutPath $titleReqPath

# --- selective data export ---
$dataIndex = New-Object System.Collections.Generic.List[object]
if ($ExportDataLists.Count -gt 0) {
  foreach ($lt in $ExportDataLists) {
    $t = [string]$lt
    if ([string]::IsNullOrWhiteSpace($t)) { continue }

    $safe = Sanitize-FileName $t
    $out = Join-Path $outDir ("data_{0}.csv" -f $safe)

    Log ("Exporting DATA: {0}" -f $t) "INFO"
    try {
      $res = Export-ListDataToCsv -ListTitle $t -OutPath $out -PageSize $DataPageSize
      $dataIndex.Add([pscustomobject]@{
          ListTitle = $t
          DataFile  = (Split-Path -Leaf $out)
          Rows      = $res.Rows
          Skipped   = $res.Skipped
        })
      Log ("DATA exported: {0} rows => {1}" -f $res.Rows, $out) "INFO"
    }
    catch {
      $dataIndex.Add([pscustomobject]@{
          ListTitle = $t
          DataFile  = (Split-Path -Leaf $out)
          Rows      = $null
          Skipped   = $false
          Error     = $_.Exception.Message
        })
      Log ("DATA export FAILED: {0} :: {1}" -f $t, $_.Exception.Message) "ERROR"
    }
  }

  $dataIndexPath = Join-Path $outDir "data_index.csv"
  $dataIndex | Export-Csv -Path $dataIndexPath -NoTypeInformation -Encoding UTF8
  Log "Wrote: $dataIndexPath" "INFO"
}
else {
  Log "No export.exportDataLists specified. Skip data export." "INFO"
}

Log "DONE export provisioning template." "INFO"
PnP-Disconnect
