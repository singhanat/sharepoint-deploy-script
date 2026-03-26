# ====================================================================
# importV2.ps1
# Created dynamically from config.json, template.xml, and data csvs
# ====================================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PSMajor { try { [int]$PSVersionTable.PSVersion.Major } catch { 5 } }
$script:UseLegacyPnP = $false
function Ensure-PnPModule {
    $major = Get-PSMajor
    if ($major -ge 7) { $script:UseLegacyPnP = $false; Import-Module PnP.PowerShell -ErrorAction Stop }
    else { $script:UseLegacyPnP = $true; Import-Module SharePointPnPPowerShellOnline -ErrorAction Stop }
}
function PnP-Connect([string]$Url) {
    if ($script:UseLegacyPnP) { Connect-PnPOnline -Url $Url -UseWebLogin }
    else { Connect-PnPOnline -Url $Url -Interactive }
}

$SiteURL = 'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX'

Write-Host 'Connecting to SharePoint: $SiteURL' -ForegroundColor Cyan
Ensure-PnPModule
PnP-Connect $SiteURL
if (-not (Get-PnPContext)) { Write-Host 'Connection failed' -ForegroundColor Red; exit }
Write-Host 'Connected.' -ForegroundColor Green

# ==================================================
# List: Exports_MMO
# ==================================================
$existingList = Get-PnPList -Identity 'Exports_MMO' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list Exports_MMO' -ForegroundColor Yellow
    Remove-PnPList -Identity 'Exports_MMO' -Force -Recycle
}
Write-Host 'Creating DocumentLibrary: Exports_MMO' -ForegroundColor Cyan
New-PnPList -Title 'Exports_MMO' -Template DocumentLibrary -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for Exports_MMO...' -ForegroundColor Cyan

Write-Host 'Updating Default View Fields for Exports_MMO...' -ForegroundColor Cyan
try {
    $viewName = if ('DocumentLibrary' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'Exports_MMO' -Identity $viewName -Fields @('FileLeafRef', 'DocIcon', 'LinkFilename', 'Modified', 'Editor') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MMO_RichTextImages
# ==================================================
$existingList = Get-PnPList -Identity 'MMO_RichTextImages' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MMO_RichTextImages' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MMO_RichTextImages' -Force -Recycle
}
Write-Host 'Creating DocumentLibrary: MMO_RichTextImages' -ForegroundColor Cyan
New-PnPList -Title 'MMO_RichTextImages' -Template DocumentLibrary -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MMO_RichTextImages...' -ForegroundColor Cyan
$fieldXml = @'
<Field Type="Note" DisplayName="Image Tags_0" Name="lcf76f155ced4ddcb4097134ff3c332f" ID="{7c51d632-27e0-f6f7-a5d6-ed67b68bd369}" ShowInViewForms="FALSE" Required="FALSE" Hidden="TRUE" CanToggleHidden="TRUE" />
'@
try { Add-PnPFieldFromXml -List 'MMO_RichTextImages' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field lcf76f155ced4ddcb4097134ff3c332f - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MMO_RichTextImages...' -ForegroundColor Cyan
try {
    $viewName = if ('DocumentLibrary' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MMO_RichTextImages' -Identity $viewName -Fields @('FileLeafRef', 'DocIcon', 'LinkFilename', 'Modified', 'Editor') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_ApprovalSteps
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_ApprovalSteps' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_ApprovalSteps' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_ApprovalSteps' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_ApprovalSteps' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_ApprovalSteps' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_ApprovalSteps...' -ForegroundColor Cyan
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="MemorandumID" Format="Dropdown" IsModern="TRUE" Name="MemorandumID" Percentage="FALSE" Required="TRUE" Title="MemorandumID" Type="Number" Unit="None" ID="{b4fcebe4-8c34-42ac-91db-241d6c68f12e}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field MemorandumID - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="ApprovalCycle" Format="Dropdown" IsModern="TRUE" Name="ApprovalCycle" Percentage="FALSE" Required="TRUE" Title="ApprovalCycle" Type="Number" Unit="None" ID="{1abb2d50-089e-480f-b703-5f102a632166}"><Default>1</Default></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field ApprovalCycle - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="ApprovalLevel" Format="Dropdown" IsModern="TRUE" Name="ApprovalLevel" Percentage="FALSE" Required="TRUE" Title="ApprovalLevel" Type="Number" Unit="None" ID="{69d158e4-7682-4cde-850d-dee775564dd5}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field ApprovalLevel - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="Approver" Format="Dropdown" IsModern="TRUE" List="UserInfo" Name="Approver" Title="Approver" Type="User" UserSelectionMode="0" UserSelectionScope="0" ID="{d172e525-99a6-49ce-a52a-483ceb7a7947}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Approver - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CustomFormatter="{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Pending&quot;]},&quot;sp-css-backgroundColor-BgCornflowerBlue sp-field-fontSizeSmall sp-css-color-CornflowerBlueFont sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Approved&quot;]},&quot;sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Returned&quot;]},&quot;sp-css-backgroundColor-BgGold sp-field-fontSizeSmall sp-css-color-GoldFont sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Rejected&quot;]},&quot;sp-css-backgroundColor-BgCoral sp-field-fontSizeSmall sp-css-color-CoralFont sp-css-color-CoralFont&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;@currentField&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Pending&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Approved&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Returned&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;Rejected&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-CoralFont&quot;,&quot;&quot;]}]}]}]}]}}}]}],&quot;templateId&quot;:&quot;BgColorChoicePill&quot;}" DisplayName="ApprovalStatus" FillInChoice="FALSE" Format="Dropdown" IsModern="TRUE" Name="ApprovalStatus" Title="ApprovalStatus" Type="Choice" ID="{21ebbe06-4179-43b3-9962-8bb5ca731c5a}"><CHOICES><CHOICE>Pending</CHOICE><CHOICE>Approved</CHOICE><CHOICE>Returned</CHOICE><CHOICE>Rejected</CHOICE></CHOICES><Default>Pending</Default></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field ApprovalStatus - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" DisplayName="Comments" Format="Dropdown" IsModern="TRUE" IsolateStyles="TRUE" Name="Comments" RichText="TRUE" RichTextMode="FullHtml" Title="Comments" Type="Note" ID="{bcd80f60-2605-4266-b90d-872d9c82423e}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Comments - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field ClientSideComponentId="00000000-0000-0000-0000-000000000000" DisplayName="ActionDate" FriendlyDisplayFormat="Disabled" Format="DateTime" Name="ActionDate" Title="ActionDate" Type="DateTime" ID="{acde0f09-858b-4a18-b853-a2128fc3297d}" Version="2" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_ApprovalSteps' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field ActionDate - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_ApprovalSteps...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_ApprovalSteps' -Identity $viewName -Fields @('MemorandumID', 'ApprovalCycle', 'ApprovalLevel', 'Approver', 'ApprovalStatus', 'Comments', 'ActionDate', 'Created', 'Author', 'ID') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_CC
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_CC' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_CC' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_CC' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_CC' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_CC' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_CC...' -ForegroundColor Cyan
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="MemorandumID" Format="Dropdown" IsModern="TRUE" Name="MemorandumID" Percentage="FALSE" Required="TRUE" Title="MemorandumID" Type="Number" Unit="None" ID="{7f669d7e-4395-4f0e-b9e6-3a4577667848}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_CC' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field MemorandumID - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="Recipient" Format="Dropdown" IsModern="TRUE" List="UserInfo" Mult="TRUE" Name="Recipient" Title="Recipient" Type="UserMulti" UserSelectionMode="0" UserSelectionScope="0" ID="{a24dd1db-5156-45bb-9cd4-6cafbcc9ba54}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_CC' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Recipient - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_CC...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_CC' -Identity $viewName -Fields @('MemorandumID', 'Recipient') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_Contracts
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_Contracts' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_Contracts' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_Contracts' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_Contracts' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_Contracts' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_Contracts...' -ForegroundColor Cyan
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="MemorandumID" Format="Dropdown" IsModern="TRUE" Name="MemorandumID" Percentage="FALSE" Required="TRUE" Title="MemorandumID" Type="Number" Unit="None" ID="{136e4782-b420-4fc0-91a5-1463a2e474ec}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Contracts' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field MemorandumID - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" ClientSideComponentId="00000000-0000-0000-0000-000000000000" DisplayName="Detail" Format="Dropdown" IsolateStyles="FALSE" Name="Detail" RichText="FALSE" RichTextMode="Compatible" Title="Detail" Type="Note" ID="{cc81dacf-7f2c-4c88-9a6f-7ba8011a7a6b}" Version="2" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Contracts' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Detail - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_Contracts...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_Contracts' -Identity $viewName -Fields @('MemorandumID', 'ID', 'Detail', 'Attachments') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_Contracts_Pictures
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_Contracts_Pictures' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_Contracts_Pictures' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_Contracts_Pictures' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_Contracts_Pictures' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_Contracts_Pictures' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_Contracts_Pictures...' -ForegroundColor Cyan
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="MemorandumID" Format="Dropdown" IsModern="TRUE" Name="MemorandumID" Percentage="FALSE" Required="TRUE" Title="MemorandumID" Type="Number" Unit="None" ID="{cc2a4801-e2ad-4d1c-99b9-2eb1949fbe68}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Contracts_Pictures' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field MemorandumID - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_Contracts_Pictures...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_Contracts_Pictures' -Identity $viewName -Fields @('MemorandumID', 'Attachments') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_DEV_MS_Department
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_DEV_MS_Department' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_DEV_MS_Department' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_DEV_MS_Department' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_DEV_MS_Department' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_DEV_MS_Department' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_DEV_MS_Department...' -ForegroundColor Cyan
$fieldXml = @'
<Field ClientSideComponentId="00000000-0000-0000-0000-000000000000" CommaSeparator="TRUE" CustomUnitOnRight="FALSE" Decimals="0" DisplayName="no" Format="Dropdown" Name="field_0" Percentage="FALSE" Title="no" Type="Number" Unit="None" ID="{8ef0df25-0ea2-48f1-9283-1fccbcbfe484}" Version="4" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_DEV_MS_Department' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field field_0 - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field Type="Text" DisplayName="code" Name="field_1" ID="{d09a273d-b6cc-4773-bba8-239a94942422}" Version="2" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_DEV_MS_Department' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field field_1 - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="active" Format="Dropdown" IsModern="TRUE" Name="active" Title="active" Type="Boolean" ID="{e2d0c7e8-03d1-404b-91c2-825ec300ea4e}"><Default>1</Default></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_DEV_MS_Department' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field active - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_DEV_MS_Department...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_DEV_MS_Department' -Identity $viewName -Fields @('LinkTitle', 'field_0', 'field_1', 'active') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

Write-Host 'Importing Data for MLN_BD_MMO_DEV_MS_Department...' -ForegroundColor Cyan
$csvData = @'
"Title","field_0","field_1","active"
"Accounting","1","ACC","True"
"Agent Business","2","AB","True"
"Automobile Industry Marketing","3","AIM","True"
"Bank and Finance Business","4","BFB","True"
"Branch-Ayutthaya","5","BR_AYT","True"
"Branch-Bang Saphan","6","BR_BSP","True"
"Branch-Buri ram","7","BR_BRR","True"
"Branch-Chachoengsao","8","BR_CCO","True"
"Branch-Chaiyaphum","9","BR_CPM","True"
"Branch-Chantaburi","10","BR_CTI","True"
"Branch-Chiang Mai","11","BR_CMI","True"
"Branch-Chiang Rai","12","BR_CRI","True"
"Branch-Chumphon","13","BR_CPN","True"
"Branch-Hat Yai","14","BR_HY","True"
"Branch-Hua-Hin","15","BR_HHN","True"
"Branch-Kala Sin","16","BR_KSN","True"
"Branch-Kamphaeng Phet","17","BR_KPP","True"
"Branch-Kanchanaburi","18","BR_KRI","True"
"Branch-Khon kaen","19","BR_KKN","True"
"Branch-Korat","20","BR_NMA","True"
"Branch-Krabi","21","BR_KBI","True"
"Branch-Lampang","22","BR_LPG","True"
"Branch-Lang Suan","23","BR_LSN","True"
"Branch-Loei","24","BR_LEI","True"
"Branch-Lop Buri","25","BR_LRI","True"
"Branch-Maha Sarakham","26","BR_MKM","True"
"Branch-Mukdahan","27","BR_MKD","True"
"Branch-Nakhon Pathom","28","BR_NPT","True"
"Branch-Nakhon Phanom","29","BR_NPM","True"
"Branch-Nakhon Sawan","30","BR_NSN","True"
"Branch-Nakhon Si Thammarat","31","BR_NST","True"
"Branch-Narathiwat","32","BR_NTW","True"
"Branch-Nong Khai","33","BR_NKI","True"
"Branch-Pattani","34","BR_PTN","True"
"Branch-Pattaya","35","BR_PTY","True"
"Branch-Phachuap Khiri Khan","36","BR_PKN","True"
"Branch-Phatthalung","37","BR_PLG","True"
"Branch-Phetchaburi","38","BR_PBI","True"
"Branch-Phitsanulok","39","BR_PLK","True"
"Branch-Phrae","40","BR_PRE","True"
"Branch-Phuket","41","BR_PKT","True"
"Branch-Prachin Buri","42","BR_PRI","True"
"Branch-Ranong","43","BR_RNG","True"
"Branch-Ratchaburi","44","BR_RBR","True"
"Branch-Rayong","45","BR_RYG","True"
"Branch-Roi Et","46","BR_RET","True"
"Branch-Sakon Nakhon","47","BR_SKN","True"
"Branch-Samui","48","BR_SMU","True"
"Branch-Samut Sakorn","49","BR_SKS","True"
"Branch-Samut Songkram","50","BR_SKG","True"
"Branch-Saraburi","51","BR_SRI","True"
"Branch-Satun","52","BR_STN","True"
"Branch-Si Racha","53","BR_SRC","True"
"Branch-Si Sa Ket","54","BR_SKI","True"
"Branch-Songkhla","55","BR_SKA","True"
"Branch-Suphan Buri","56","BR_SPB","True"
"Branch-Surat Thani","57","BR_SNI","True"
"Branch-Surin","58","BR_SRN","True"
"Branch-Takua Pa","59","BR_TKP","True"
"Branch-Thung Song","60","BR_TSG","True"
"Branch-Trang","61","BR_TRG","True"
"Branch-Ubon Ratchathani","62","BR_UBN","True"
"Branch-Udon Thani","63","BR_UDN","True"
"Branch-Wieng Sa","64","BR_WSA","True"
"Branch-Yala","65","BR_YLA","True"
"Branch Operation","66","BROPS","True"
"Business Development","67","BD","True"
"Commercial Underwriting","68","CUW","True"
"Compliance & Legal, and Corporate Secretary","69","CLCS","True"
"Corporate Management","70","CM","True"
"Dealer Business 1","71","DB1","True"
"Dealer Business 2","72","DB2","True"
"Finance","73","FIN","True"
"General Affairs","74","GA","True"
"Human Resources","75","HR","True"
"Information Technology","76","IT","True"
"Internal Audit","77","IA","True"
"International Broker","78","IB","True"
"Investment","79","INV","True"
"Local Broker","80","LB","True"
"Management Office","81","MOF","True"
"Motor Claims","82","MC","True"
"Multinational Marketing","83","MNM","True"
"Non-motor Claims","84","NMC","True"
"Production","85","PRO","True"
"Recovery","86","RCV","True"
"Retail Underwriting","87","RUW","True"
"Risk Management","88","RM","True"
"TM Claim Services","89","TMCS","True"
"All Department",,,"True"
'@
$items = $csvData | ConvertFrom-Csv
foreach ($itm in $items) {
    $values = @{}
    if ($null -ne $itm.'active') { $values['active'] = $itm.'active' }
    if ($null -ne $itm.'field_0') { $values['field_0'] = $itm.'field_0' }
    if ($null -ne $itm.'field_1') { $values['field_1'] = $itm.'field_1' }
    if ($null -ne $itm.'Title') { $values['Title'] = $itm.'Title' }
    if ($values.Count -gt 0) {
        try { Add-PnPListItem -List 'MLN_BD_MMO_DEV_MS_Department' -Values $values | Out-Null } catch { Write-Host "Failed adding item in MLN_BD_MMO_DEV_MS_Department" -ForegroundColor Yellow }
    }
}

# ==================================================
# List: MLN_BD_MMO_MailTemplates
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_MailTemplates' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_MailTemplates' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_MailTemplates' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_MailTemplates' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_MailTemplates' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_MailTemplates...' -ForegroundColor Cyan
$fieldXml = @'
<Field DisplayName="Subject" Format="Dropdown" IsModern="TRUE" MaxLength="255" Name="Subject" Title="Subject" Type="Text" ID="{06271359-fea2-4a2f-bc08-54a246895f23}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_MailTemplates' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Subject - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" DisplayName="Body" Format="Dropdown" IsModern="TRUE" IsolateStyles="TRUE" Name="Body" RichText="TRUE" RichTextMode="FullHtml" Title="Body" Type="Note" ID="{d2979909-c9db-4c93-a402-b0b08690a2b1}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_MailTemplates' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Body - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_MailTemplates...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_MailTemplates' -Identity $viewName -Fields @('LinkTitle', 'Subject', 'Body') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

Write-Host 'Importing Data for MLN_BD_MMO_MailTemplates...' -ForegroundColor Cyan
$csvData = @'
"Title","Subject","Body"
"NewRequestForApproval","Memorandum for Approval: [ProjectName]","<div class=""ExternalClassAE8F859BE99C4FAD866889472A3C2E89""><div class=""ExternalClass60764E02DFD54ACD999DE6BFC5DBF8C2"" style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;"">Dear<span>&#160;</span><b>[ApproverName],</b></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><div style=""margin&#58;14.6667px 0px;"">A new memorandum titled<span>&#160;</span><b>&quot;[ProjectName]&quot;&#160;requires your approval.</b><br></div>Please review and take action at the link below&#58;<div style=""margin&#58;14.6667px 0px;""><strong style=""box-sizing&#58;border-box;color&#58;rgb(226, 226, 229);font-family&#58;Inter, sans-serif;font-size&#58;14px;background-color&#58;rgb(25, 25, 25);transition&#58;none  !important;""></strong><span style=""color&#58;rgb(12, 100, 192);""><span style=""color&#58;rgb(23, 78, 134);font-family&#58;Calibri, Arial, Helvetica, sans-serif;background-color&#58;rgb(255, 255, 255);display&#58;inline !important;"">[LinkToItem]</span>​</span><br></div></div><div class=""ExternalClass60764E02DFD54ACD999DE6BFC5DBF8C2""><div><br></div><span></span></div></div>"
"RequestApproved","Approved: Memorandum [ProjectName]","<div class=""ExternalClass94589B7D52AF4C0CA5FBD5302C7669E9""><div style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;"">Dear<b>&#160;[RequesterName],</b></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><div style=""margin&#58;14.6667px 0px;"">This is to inform you that your memorandum<b><span>&#160;</span>&quot;[ProjectName]&quot; has been fully<span>&#160;</span></b><span style=""color&#58;rgb(12, 136, 42);""><b>Approved.</b></span><span style=""color&#58;rgb(12, 136, 42);"">​<div style=""margin&#58;14.6667px 0px;""><span style=""color&#58;rgb(0, 0, 0);"">You can view the full report and approval history at the link below&#58;</span><div style=""margin&#58;14.6667px 0px;""><span style=""color&#58;rgb(12, 100, 192);"">[LinkToItem]</span></div></div></span></div></div></div>"
"RequestRejected","Memorandum [ProjectName]","<div class=""ExternalClass8E3D713057F14165A9526434E68A9E77""><div style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;"">Dear<b><span>&#160;</span>[RequesterName],</b></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><div style=""margin&#58;14.6667px 0px;"">This is to inform you that your memorandum<span>&#160;</span><span style=""font-weight&#58;bold;"">&quot;[ProjectName]&quot;&#160;has been<span>&#160;</span></span><span style=""font-weight&#58;bold;color&#58;rgb(200, 38, 19);"">Rejected.</span><span style=""font-weight&#58;bold;""><br></span></div>You can view the full report and approval history at the link below&#58;<div style=""margin&#58;14.6667px 0px;""><span style=""color&#58;rgb(12, 100, 192);"">[LinkToItem]</span><br></div></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><br style=""font-size&#58;14.6667px;background-color&#58;rgb(255, 255, 255);""></div><div style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span><br></div></div>"
"RequestSentBack","Action Required: Memorandum [ProjectName] has been sent back","<div class=""ExternalClassDB7DB601BB954F8E8F8716179FB9BDE3""><div style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;"">Dear<span>&#160;</span><b>[RequesterName],</b></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><div style=""margin&#58;14.6667px 0px;"">Your memorandum<b><span>&#160;</span>&quot;[ProjectName]&quot;<span>&#160;</span></b>has been<span style=""color&#58;rgb(0, 0, 0);""><b><span>&#160;</span></b></span><span style=""color&#58;rgb(0, 0, 0);""><b>Sent Back<span>&#160;</span></b></span>for revision.<div style=""margin&#58;14.6667px 0px;"">Approver's comment&#58;<b><span>&#160;</span></b><span style=""color&#58;rgb(23, 78, 134);""><b>[ApproverComments]</b></span><p><br></p><div style=""margin&#58;14.6667px 0px;"">Please review the comments, update the request, and resubmit for approval at the link below&#58;<div style=""margin&#58;14.6667px 0px;""><span style=""color&#58;rgb(12, 100, 192);"">[LinkToItem]</span></div></div></div></div></div></div>"
"RequestDeleted","FYI: Memorandum ""[ProjectName]"" has been deleted","<div class=""ExternalClassA3BD620A76E24180B17F2DAB0CC76A17""><div style=""font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;color&#58;rgb(50, 49, 48);background-color&#58;transparent;""><span></span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><span>Dear All,</span></div><div style=""margin&#58;0px;font-family&#58;Calibri, Arial, Helvetica, sans-serif;font-size&#58;11pt;""><div style=""margin&#58;0px;"">This is to inform you that the memorandum titled<b><span>&#160;</span>&quot;[ProjectName]&quot;&#160;</b>(Doc No&#58;<span>&#160;</span><b>[DocNo]</b>) has been<span>&#160;</span><span style=""color&#58;rgb(200, 38, 19);"">deleted by</span><span>&#160;</span><b>[DeletedBy]</b>.<br></div><div style=""margin&#58;0px;""><p><span>This request has been removed from the approval process. No further action is required from your side.</span></p></div></div></div>"
'@
$items = $csvData | ConvertFrom-Csv
foreach ($itm in $items) {
    $values = @{}
    if ($null -ne $itm.'Body') { $values['Body'] = $itm.'Body' }
    if ($null -ne $itm.'Subject') { $values['Subject'] = $itm.'Subject' }
    if ($null -ne $itm.'Title') { $values['Title'] = $itm.'Title' }
    if ($values.Count -gt 0) {
        try { Add-PnPListItem -List 'MLN_BD_MMO_MailTemplates' -Values $values | Out-Null } catch { Write-Host "Failed adding item in MLN_BD_MMO_MailTemplates" -ForegroundColor Yellow }
    }
}

# ==================================================
# List: MLN_BD_MMO_Memorandums
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_Memorandums' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_Memorandums' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_Memorandums' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_Memorandums' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_Memorandums' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_Memorandums...' -ForegroundColor Cyan
$fieldXml = @'
<Field DisplayName="DocNo" Format="Dropdown" IsModern="TRUE" MaxLength="255" Name="DocNo" Title="DocNo" Type="Text" ID="{c9636614-d39c-42eb-a5c0-532b0461af62}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field DocNo - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="Department" Format="Dropdown" IsModern="TRUE" MaxLength="255" Name="Department" Required="TRUE" Title="Department" Type="Text" ID="{a4d9a37e-6387-4ddc-933d-10a1cb4115e1}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Department - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" DisplayName="Objective" Format="Dropdown" IsModern="TRUE" IsolateStyles="TRUE" Name="Objective" RichText="TRUE" RichTextMode="FullHtml" Title="Objective" Type="Note" ID="{85dc6da3-1b31-410b-888b-813d835a2977}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Objective - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field ClientSideComponentId="00000000-0000-0000-0000-000000000000" CustomFormatter="{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Draft&quot;]},&quot;sp-css-backgroundColor-BgCornflowerBlue sp-css-color-CornflowerBlueFont sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Waiting&quot;]},&quot;sp-css-backgroundColor-BgMintGreen sp-css-color-MintGreenFont sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Returned&quot;]},&quot;sp-css-backgroundColor-BgGold sp-css-color-GoldFont sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Approved&quot;]},&quot;sp-css-backgroundColor-BgCoral sp-css-color-CoralFont sp-css-color-CoralFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Rejected&quot;]},&quot;sp-css-backgroundColor-BgDustRose sp-css-color-DustRoseFont sp-css-color-DustRoseFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Deleted&quot;]},&quot;sp-css-backgroundColor-BgCyan sp-css-color-CyanFont sp-css-color-CyanFont&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;[$RequestStatus]&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Draft&quot;]},&quot;sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Waiting&quot;]},&quot;sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Returned&quot;]},&quot;sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Approved&quot;]},&quot;sp-css-color-CoralFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Rejected&quot;]},&quot;sp-css-color-DustRoseFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$RequestStatus]&quot;,&quot;Deleted&quot;]},&quot;sp-css-color-CyanFont&quot;,&quot;&quot;]}]}]}]}]}]}]}}}]}],&quot;templateId&quot;:&quot;BgColorChoicePill&quot;}" DisplayName="RequestStatus" FillInChoice="FALSE" Format="Dropdown" Name="RequestStatus" Title="RequestStatus" Type="Choice" ID="{56c938a7-e589-4a70-8335-002ce579d963}" Version="2"><CHOICES><CHOICE>Draft</CHOICE><CHOICE>Waiting</CHOICE><CHOICE>Returned</CHOICE><CHOICE>Approved</CHOICE><CHOICE>Rejected</CHOICE><CHOICE>Deleted</CHOICE></CHOICES><Default>Draft</Default></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field RequestStatus - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field ClientSideComponentId="00000000-0000-0000-0000-000000000000" CustomFormatter="{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;Organize&quot;]},&quot;sp-css-backgroundColor-BgCornflowerBlue sp-css-color-CornflowerBlueFont sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;Premium&quot;]},&quot;sp-css-backgroundColor-BgMintGreen sp-css-color-MintGreenFont sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;ทดสอบ&quot;]},&quot;sp-css-backgroundColor-BgGold sp-css-color-GoldFont sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;สิ่งพิมพ์&quot;]},&quot;sp-css-backgroundColor-BgCoral sp-css-color-CoralFont sp-css-color-CoralFont&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;[$ProcurementType]&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;Organize&quot;]},&quot;sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;Premium&quot;]},&quot;sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;ทดสอบ&quot;]},&quot;sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;&quot;]},&quot;&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;[$ProcurementType]&quot;,&quot;สิ่งพิมพ์&quot;]},&quot;sp-css-color-CoralFont&quot;,&quot;&quot;]}]}]}]}]}}}]}],&quot;templateId&quot;:&quot;BgColorChoicePill&quot;}" DisplayName="ProcurementType" FillInChoice="TRUE" Format="Dropdown" Name="ProcurementType" Required="TRUE" Title="ProcurementType" Type="Choice" ID="{e48ddc91-eceb-47ca-bf56-eff7f485dd6e}" Version="2"><CHOICES><CHOICE>Organize</CHOICE><CHOICE>Premium</CHOICE><CHOICE>ทดสอบ</CHOICE><CHOICE>สิ่งพิมพ์</CHOICE></CHOICES></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field ProcurementType - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" Decimals="2" DisplayName="TotalCostNoVAT" Format="Dropdown" IsModern="TRUE" Name="TotalCostNoVAT" Percentage="FALSE" Title="TotalCostNoVAT" Type="Number" Unit="None" ID="{d59b4de3-5445-4c4d-ab69-2e2d247d2c7f}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field TotalCostNoVAT - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" Decimals="0" DisplayName="Units" Format="Dropdown" IsModern="TRUE" Name="Units" Percentage="FALSE" Title="Units" Type="Number" Unit="None" ID="{5ed560f3-8808-4014-8d5b-dc3316a77f88}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Units - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" ClientSideComponentId="00000000-0000-0000-0000-000000000000" DisplayName="Detail" Format="Dropdown" IsolateStyles="FALSE" Name="Detail" RichText="FALSE" RichTextMode="Compatible" Title="Detail" Type="Note" ID="{59060530-9af0-4642-9a74-1848f46ef677}" Version="2" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Detail - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="CurrentApprover" Format="Dropdown" IsModern="TRUE" List="UserInfo" Name="CurrentApprover" Title="CurrentApprover" Type="User" UserSelectionMode="0" UserSelectionScope="0" ID="{906b2c85-125e-43c7-8712-c383b5dd46e6}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field CurrentApprover - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field ClientSideComponentId="00000000-0000-0000-0000-000000000000" CommaSeparator="TRUE" CustomUnitOnRight="TRUE" Decimals="2" DisplayName="TotalCostWithVAT" Format="Dropdown" Name="TotalCost" Percentage="FALSE" Title="TotalCostWithVAT" Type="Number" Unit="None" ID="{444bb6a2-9160-4984-b3af-2ff7e3a993ac}" Version="2" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field TotalCost - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field DisplayName="IsDelete" Format="Dropdown" IsModern="TRUE" Name="IsDelete" Title="IsDelete" Type="Boolean" ID="{7ce8b2ed-924f-4ac1-82f3-e7b58a51f545}"><Default>0</Default></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field IsDelete - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_Memorandums...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_Memorandums' -Identity $viewName -Fields @('LinkTitle', 'DocNo', 'Department', 'Objective', 'RequestStatus', 'ProcurementType', 'TotalCostNoVAT', 'Units', 'Detail', 'CurrentApprover', 'Modified', 'Created', 'Author', 'Editor', 'Attachments', 'TotalCost', 'ID', 'IsDelete') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_Memorandums_Pictures
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_Memorandums_Pictures' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_Memorandums_Pictures' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_Memorandums_Pictures' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_Memorandums_Pictures' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_Memorandums_Pictures' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_Memorandums_Pictures...' -ForegroundColor Cyan
$fieldXml = @'
<Field CommaSeparator="TRUE" CustomUnitOnRight="TRUE" DisplayName="MemorandumID" Format="Dropdown" IsModern="TRUE" Name="MemorandumID" Percentage="FALSE" Required="TRUE" Title="MemorandumID" Type="Number" Unit="None" ID="{037e6d98-c3eb-443c-9fd7-d28a4e34c300}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_Memorandums_Pictures' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field MemorandumID - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_Memorandums_Pictures...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_Memorandums_Pictures' -Identity $viewName -Fields @('MemorandumID', 'Attachments') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

# ==================================================
# List: MLN_BD_MMO_MetaData
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_MetaData' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_MetaData' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_MetaData' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_MetaData' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_MetaData' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_MetaData...' -ForegroundColor Cyan
$fieldXml = @'
<Field DisplayName="Key" Format="Dropdown" IsModern="TRUE" MaxLength="255" Name="Key" Title="Key" Type="Text" ID="{edfd9397-17c0-477c-b2da-77f30bd785ef}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_MetaData' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Key - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field AppendOnly="FALSE" DisplayName="Value" Format="Dropdown" IsModern="TRUE" IsolateStyles="FALSE" Name="Value" RichText="FALSE" RichTextMode="Compatible" Title="Value" Type="Note" ID="{a3026f5e-074b-4879-833b-61eaa637d22d}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_MetaData' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Value - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_MetaData...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_MetaData' -Identity $viewName -Fields @('LinkTitle', 'Key', 'Value') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }

Write-Host 'Importing Data for MLN_BD_MMO_MetaData...' -ForegroundColor Cyan
$csvData = @'
"Title","Key","Value"
,"AppLink","https://apps.powerapps.com/play/e/17f72c25-782e-ed8d-a19e-747a92631c90/a/1d6e736a-34a0-4427-ab0b-e4b0b8d705a9?tenantId=9e3320d5-0a75-4623-8d30-807c36031583&hint=cafd0554-bd55-4a4a-a075-7428bca7c845&sourcetime=1769489382866&ItemID=__ITEM_ID__"
'@
$items = $csvData | ConvertFrom-Csv
foreach ($itm in $items) {
    $values = @{}
    if ($null -ne $itm.'Key') { $values['Key'] = $itm.'Key' }
    if ($null -ne $itm.'Title') { $values['Title'] = $itm.'Title' }
    if ($null -ne $itm.'Value') { $values['Value'] = $itm.'Value' }
    if ($values.Count -gt 0) {
        try { Add-PnPListItem -List 'MLN_BD_MMO_MetaData' -Values $values | Out-Null } catch { Write-Host "Failed adding item in MLN_BD_MMO_MetaData" -ForegroundColor Yellow }
    }
}

# ==================================================
# List: MLN_BD_MMO_UserRoles
# ==================================================
$existingList = Get-PnPList -Identity 'MLN_BD_MMO_UserRoles' -ErrorAction SilentlyContinue
if ($existingList) {
    Write-Host 'Deleting existing list MLN_BD_MMO_UserRoles' -ForegroundColor Yellow
    Remove-PnPList -Identity 'MLN_BD_MMO_UserRoles' -Force -Recycle
}
Write-Host 'Creating GenericList: MLN_BD_MMO_UserRoles' -ForegroundColor Cyan
New-PnPList -Title 'MLN_BD_MMO_UserRoles' -Template GenericList -OnQuickLaunch:$false | Out-Null

Write-Host 'Adding fields for MLN_BD_MMO_UserRoles...' -ForegroundColor Cyan
$fieldXml = @'
<Field DisplayName="User" Format="Dropdown" IsModern="TRUE" List="UserInfo" Name="User" Title="User" Type="User" UserSelectionMode="0" UserSelectionScope="0" ID="{69f7080d-63a0-4d32-8054-ab6a3895980c}" />
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_UserRoles' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field User - $($_.Exception.Message)" -ForegroundColor Yellow }
$fieldXml = @'
<Field CustomFormatter="{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;flex-wrap&quot;:&quot;wrap&quot;,&quot;display&quot;:&quot;flex&quot;},&quot;children&quot;:[{&quot;elmType&quot;:&quot;div&quot;,&quot;style&quot;:{&quot;box-sizing&quot;:&quot;border-box&quot;,&quot;padding&quot;:&quot;4px 8px 5px 8px&quot;,&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;display&quot;:&quot;flex&quot;,&quot;border-radius&quot;:&quot;16px&quot;,&quot;height&quot;:&quot;24px&quot;,&quot;align-items&quot;:&quot;center&quot;,&quot;white-space&quot;:&quot;nowrap&quot;,&quot;margin&quot;:&quot;4px 4px 4px 4px&quot;},&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_REQUESTER&quot;]},&quot;sp-css-backgroundColor-BgCornflowerBlue sp-field-fontSizeSmall sp-css-color-CornflowerBlueFont sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_EDITOR&quot;]},&quot;sp-css-backgroundColor-BgMintGreen sp-field-fontSizeSmall sp-css-color-MintGreenFont sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_ADMIN&quot;]},&quot;sp-css-backgroundColor-BgGold sp-field-fontSizeSmall sp-css-color-GoldFont sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;sp-field-borderAllRegular sp-field-borderAllSolid sp-css-borderColor-neutralSecondary&quot;]}]}]}]}},&quot;children&quot;:[{&quot;elmType&quot;:&quot;span&quot;,&quot;style&quot;:{&quot;overflow&quot;:&quot;hidden&quot;,&quot;text-overflow&quot;:&quot;ellipsis&quot;,&quot;padding&quot;:&quot;0 3px&quot;},&quot;txtContent&quot;:&quot;@currentField&quot;,&quot;attributes&quot;:{&quot;class&quot;:{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_REQUESTER&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-CornflowerBlueFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_EDITOR&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-MintGreenFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;ROLE_ADMIN&quot;]},&quot;sp-field-fontSizeSmall sp-css-color-GoldFont&quot;,{&quot;operator&quot;:&quot;:&quot;,&quot;operands&quot;:[{&quot;operator&quot;:&quot;==&quot;,&quot;operands&quot;:[&quot;@currentField&quot;,&quot;&quot;]},&quot;&quot;,&quot;&quot;]}]}]}]}}}]}],&quot;templateId&quot;:&quot;BgColorChoicePill&quot;}" DisplayName="Role" FillInChoice="FALSE" Format="Dropdown" IsModern="TRUE" Name="Role" Title="Role" Type="Choice" ID="{a41bdc49-5a6e-4002-8493-32820555de40}"><CHOICES><CHOICE>ROLE_REQUESTER</CHOICE><CHOICE>ROLE_EDITOR</CHOICE><CHOICE>ROLE_ADMIN</CHOICE></CHOICES></Field>
'@
try { Add-PnPFieldFromXml -List 'MLN_BD_MMO_UserRoles' -FieldXml $fieldXml | Out-Null } catch { Write-Host "Error adding field Role - $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host 'Updating Default View Fields for MLN_BD_MMO_UserRoles...' -ForegroundColor Cyan
try {
    $viewName = if ('GenericList' -eq 'DocumentLibrary') { 'All Documents' } else { 'All Items' }
    Set-PnPView -List 'MLN_BD_MMO_UserRoles' -Identity $viewName -Fields @('User', 'Role') -ErrorAction Stop | Out-Null
    Write-Host 'View fields updated.' -ForegroundColor DarkGreen
}
catch { Write-Host 'Failed to update view fields: ' + $_.Exception.Message -ForegroundColor Yellow }


