param(
    [string]$ConfigPath = "c:\Work\tokio-marine\config.json",
    [string]$OutScript = "c:\Work\tokio-marine\importV2.ps1",
    [string]$ExportDir = "c:\Work\tokio-marine\MSG-PRJ-UAT_PowerApp"
)

$ErrorActionPreference = "Stop"

# 1. Read config
$config = Get-Content -Raw -Path $ConfigPath | ConvertFrom-Json
$targetSiteUrl = $config.import.targetSiteUrl
if (-not $targetSiteUrl) { throw "No target site url found" }

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $scriptDir) { $scriptDir = "." }

$sourceFolder = $config.import.sourceFolder
$ExportDir = Join-Path $scriptDir $sourceFolder
if (-not (Test-Path $ExportDir)) {
    throw "Source folder not found: $ExportDir"
}

$sb = [System.Text.StringBuilder]::new()

$sb.AppendLine("# ====================================================================")
$sb.AppendLine("# importV2.ps1")
$sb.AppendLine("# Created dynamically from config.json, template.xml, and data csvs")
$sb.AppendLine("# ====================================================================")
$sb.AppendLine("")
$sb.AppendLine("Set-StrictMode -Version Latest")
$sb.AppendLine("`$ErrorActionPreference = 'Stop'")
$sb.AppendLine("")
$sb.AppendLine("function Get-PSMajor { try { [int]`$PSVersionTable.PSVersion.Major } catch { 5 } }")
$sb.AppendLine("`$script:UseLegacyPnP = `$false")
$sb.AppendLine("function Ensure-PnPModule {")
$sb.AppendLine("  `$major = Get-PSMajor")
$sb.AppendLine("  if (`$major -ge 7) { `$script:UseLegacyPnP = `$false; Import-Module PnP.PowerShell -ErrorAction Stop }")
$sb.AppendLine("  else { `$script:UseLegacyPnP = `$true; Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop }")
$sb.AppendLine("}")
$sb.AppendLine("function PnP-Connect([string]`$Url) {")
$sb.AppendLine("  if (`$script:UseLegacyPnP) { Connect-PnPOnline -Url `$Url -UseWebLogin }")
$sb.AppendLine("  else { Connect-PnPOnline -Url `$Url -Interactive }")
$sb.AppendLine("}")
$sb.AppendLine("")
$sb.AppendLine("`$SiteURL = '$targetSiteUrl'")
$sb.AppendLine("")
$sb.AppendLine("Write-Host 'Connecting to SharePoint: `$SiteURL' -ForegroundColor Cyan")
$sb.AppendLine("Ensure-PnPModule")
$sb.AppendLine("PnP-Connect `$SiteURL")
$sb.AppendLine("if (-not (Get-PnPContext)) { Write-Host 'Connection failed' -ForegroundColor Red; exit }")
$sb.AppendLine("Write-Host 'Connected.' -ForegroundColor Green")
$sb.AppendLine("")

# 2. Add Groups and Site Permissions (Disabled as requested)
<#
if ($config.import.permissions.manageSiteGroups) {
    $sb.AppendLine("# --- Site Groups ---")
    foreach ($grp in $config.import.permissions.manageSiteGroups) {
        $sb.AppendLine("Write-Host 'Ensuring Site Group: $($grp.name)' -ForegroundColor Cyan")
        $sb.AppendLine("`$grpObj = Get-PnPGroup -Identity '$($grp.name)' -ErrorAction SilentlyContinue")
        $sb.AppendLine("if (-not `$grpObj) {")
        $sb.AppendLine("    Write-Host 'Creating group: $($grp.name)...' -ForegroundColor Green")
        $sb.AppendLine("    New-PnPGroup -Title '$($grp.name)' -ErrorAction Stop | Out-Null")
        $sb.AppendLine("} else { Write-Host 'Group $($grp.name) already exists' -ForegroundColor DarkGray }")
    }
    $sb.AppendLine("")
}
#>

# Read lists_index
$listsIndexPath = Join-Path $ExportDir "lists_index.csv"
$listsIndex = @{}
if (Test-Path $listsIndexPath) {
    Import-Csv $listsIndexPath | ForEach-Object { $listsIndex[$_.Title] = $_.BaseTemplate }
}

# Parse XML for schemas
$xmlContent = Get-Content -Raw -Path (Join-Path $ExportDir "template.xml") -Encoding UTF8
$xml = [xml]$xmlContent

$nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$nsmgr.AddNamespace("pnp", "http://schemas.dev.office.com/PnP/2020/02/ProvisioningSchema")

# 3. Provision Lists
$importLists = @($config.import.importLists)
$importDataLists = @($config.import.importDataLists)

foreach ($listTitle in $importLists) {
    if ([string]::IsNullOrWhiteSpace($listTitle)) { continue }
    $sb.AppendLine("# " + ("=" * 50))
    $sb.AppendLine("# List: $listTitle")
    $sb.AppendLine("# " + ("=" * 50))
    
    $baseTemplate = $listsIndex[$listTitle]
    if (-not $baseTemplate) { $baseTemplate = "100" } # Default to GenericList
    
    $templateType = if ($baseTemplate -eq "101") { "DocumentLibrary" } else { "GenericList" }
    
    # Check if list exists, delete if configured
    if ($config.import.deleteIfExists) {
        $sb.AppendLine("`$existingList = Get-PnPList -Identity '$listTitle' -ErrorAction SilentlyContinue")
        $sb.AppendLine("if (`$existingList) {")
        $sb.AppendLine("    Write-Host 'Deleting existing list $listTitle' -ForegroundColor Yellow")
        $sb.AppendLine("    Remove-PnPList -Identity '$listTitle' -Force -Recycle")
        $sb.AppendLine("}")
    }
    
    # Create List
    $sb.AppendLine("Write-Host 'Creating $($templateType): $listTitle' -ForegroundColor Cyan")
    $onQuickLaunch = "`$false"
    $sb.AppendLine("New-PnPList -Title '$listTitle' -Template $templateType -OnQuickLaunch:$onQuickLaunch | Out-Null")
    $sb.AppendLine("")
    
    # Add Fields
    $sb.AppendLine("Write-Host 'Adding fields for $listTitle...' -ForegroundColor Cyan")
    $listNode = $xml.SelectSingleNode("//pnp:ListInstance[@Title='$listTitle']", $nsmgr)
    
    $fields = @()
    if ($listNode) {
        $fieldNodes = $listNode.SelectNodes("pnp:Fields/Field", $nsmgr)
        foreach ($node in $fieldNodes) {
            $fName = $node.GetAttribute("Name")
            # Clear RowOrdinal and SourceID as it may block import
            $node.RemoveAttribute("RowOrdinal")
            $node.RemoveAttribute("SourceID")
            $node.RemoveAttribute("ColName")
            $node.RemoveAttribute("StaticName")
            
            $sb.AppendLine("`$fieldXml = @'")
            $sb.AppendLine($node.OuterXml)
            $sb.AppendLine("'@")
            $sb.AppendLine("try { Add-PnPFieldFromXml -List '$listTitle' -FieldXml `$fieldXml | Out-Null } catch { Write-Host `"Error adding field $fName - `$(`$_.Exception.Message)`" -ForegroundColor Yellow }")
            $fields += $fName
        }
        
    }
    $sb.AppendLine("")
    
    # Views Update ViewFields
    $defaultViewNode = $listNode.SelectSingleNode("*[local-name()='Views']/*[local-name()='View'][@DefaultView='TRUE']", $nsmgr)
    if ($defaultViewNode) {
        $vFields = @()
        # FieldRef nodes might not have a namespace prefix in the XML
        $vfNodes = $defaultViewNode.SelectNodes(".//*[local-name()='FieldRef']")
        foreach ($vf in $vfNodes) {
            $vfName = $vf.GetAttribute("Name")
            if ($vfName) { $vFields += $vfName }
        }
        if ($vFields.Count -gt 0) {
            $sb.AppendLine("Write-Host 'Updating Default View Fields for $listTitle...' -ForegroundColor Cyan")
            $vfsJoined = ($vFields | ForEach-Object { "'$_'" }) -join ","
            $sb.AppendLine("try {")
            $sb.AppendLine("    `$viewName = if ('$templateType' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }")
            $sb.AppendLine("    Set-PnPView -List '$listTitle' -Identity `$viewName -Fields @($vfsJoined) -ErrorAction Stop | Out-Null")
            $sb.AppendLine("    Write-Host 'View fields updated.' -ForegroundColor DarkGreen")
            $sb.AppendLine("} catch { Write-Host 'Failed to update view fields: ' + `$_.Exception.Message -ForegroundColor Yellow }")
        }
    }
    $sb.AppendLine("")
    
    # Data insertion
    if ($importDataLists -contains $listTitle) {
        $dataCsvPath = Join-Path $ExportDir "data_$listTitle.csv"
        if (Test-Path $dataCsvPath) {
            $sb.AppendLine("Write-Host 'Importing Data for $listTitle...' -ForegroundColor Cyan")
            $csvRaw = Get-Content -Raw $dataCsvPath
            $sb.AppendLine("`$csvData = @'")
            $sb.AppendLine($csvRaw.Trim())
            $sb.AppendLine("'@")
            $sb.AppendLine("`$items = `$csvData | ConvertFrom-Csv")
            $sb.AppendLine("foreach (`$itm in `$items) {")
            $sb.AppendLine("    `$values = @{}")
            
            # Identify columns dynamically
            $headers = (Import-Csv $dataCsvPath | Get-Member -MemberType NoteProperty).Name
            foreach ($h in $headers) {
                # Skip read-only columns for import unless needed
                if ($h -match "^ID$|^Author$|^Editor$|^Created$|^Modified$") { continue }
                $sb.AppendLine("    if (`$null -ne `$itm.'$h') { `$values['$h'] = `$itm.'$h' }")
            }
            $sb.AppendLine("    if (`$values.Count -gt 0) {")
            $sb.AppendLine("        try { Add-PnPListItem -List '$listTitle' -Values `$values | Out-Null } catch { Write-Host `"Failed adding item in $listTitle`" -ForegroundColor Yellow }")
            $sb.AppendLine("    }")
            $sb.AppendLine("}")
            $sb.AppendLine("")
        }
    }
    
    # Permissions Settings (Disabled as requested)
    <#
    $sb.AppendLine("Write-Host 'Configuring Permissions for $listTitle...' -ForegroundColor Cyan")
    
    # default list permissions
    $applyDef = $config.import.permissions.defaultListPermission.applyDefault
    $stopInhDef = $config.import.permissions.defaultListPermission.stopInheriting
    $copyRoleDef = $config.import.permissions.defaultListPermission.copyRoleAssignments
    $clearDef = $config.import.permissions.defaultListPermission.clearExisting
    
    # overrides
    $override = $config.import.permissions.manageListPermissions | Where-Object { $_.listTitle -eq $listTitle }
    
    $stopInh = $stopInhDef; $copyRole = $copyRoleDef; $clear = $clearDef; $targetAssigns = $null
    
    if ($applyDef) { $targetAssigns = $config.import.permissions.defaultListPermission.assignments }
    if ($override) {
        $stopInh = $override.stopInheriting
        $copyRole = $override.copyRoleAssignments
        $clear = $override.clearExisting
        $targetAssigns = $override.assignments
    }
    
    # Clear / Stop inheriting 
    if ($stopInh -eq $true -or $override) {
        $copyRoleStr = if ($copyRole) { "`$true" } else { "`$false" }
        $clearStr = if ($clear) { "`$true" } else { "`$false" }
        $sb.AppendLine("if (`$script:UseLegacyPnP) {")
        $sb.AppendLine("    Set-PnPList -Identity '$listTitle' -BreakRoleInheritance -CopyRoleAssignments:$copyRoleStr -ClearSubscopes:$clearStr")
        $sb.AppendLine("} else {")
        $sb.AppendLine("    Set-PnPListPermission -Identity '$listTitle' -BreakRoleInheritance -CopyRoleAssignments:$copyRoleStr -ClearExisting:$clearStr")
        $sb.AppendLine("}")
    }
    
    if ($targetAssigns) {
        foreach ($asgn in $targetAssigns) {
            # Note: For groups PnP expects "-Group". For users, "-User". Here assuming group.
            $sb.AppendLine("Set-PnPListPermission -Identity '$listTitle' -Group '$($asgn.groupName)' -AddRole '$($asgn.permission)'")
        }
    }
    
    $sb.AppendLine("")
    #>
}

Out-File -FilePath $OutScript -InputObject $sb.ToString() -Encoding UTF8
Write-Host "Generated $OutScript successfully." -ForegroundColor Green
