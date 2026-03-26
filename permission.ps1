# ==============================
# Create Group + Apply Permissions to Multiple Lists
# ==============================

$SiteURL = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
Connect-PnPOnline -Url $SiteURL -UseWebLogin

# Group to create and assign
$GroupName = "MMO-User"

# Create group if not exists
$existingGroup = Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue
if (-not $existingGroup) {
    Write-Host "Creating group: $GroupName" -ForegroundColor Cyan
    New-PnPGroup -Title $GroupName -Description "Permission group for sub list"
}
else {
    Write-Host "Group already exists: $GroupName" -ForegroundColor Yellow
}

# List names to loop
$ListNames = @(
    "Exports_MMO",
    "MMO_RichTextImages",
    "MLN_BD_MMO_ApprovalSteps",
    "MLN_BD_MMO_CC",
    "MLN_BD_MMO_Contracts",
    "MLN_BD_MMO_Contracts_Pictures",
    "MLN_BD_MMO_DEV_MS_Department",
    "MLN_BD_MMO_MailTemplates",
    "MLN_BD_MMO_Memorandums",
    "MLN_BD_MMO_Memorandums_Pictures",
    "MLN_BD_MMO_MetaData",
    "MLN_BD_MMO_UserRoles"
)

foreach ($ListName in $ListNames) {

    Write-Host "`nProcessing list: $ListName" -ForegroundColor Cyan

    # Get list
    $list = Get-PnPList -Identity $ListName -ErrorAction Stop

    # Break inheritance only if needed
    if (-not $list.HasUniqueRoleAssignments) {
        Write-Host "Breaking permission inheritance for $ListName"
        $list.BreakRoleInheritance($true, $true)
        Invoke-PnPQuery
    }

    # Assign permission
    Write-Host "Assigning Edit permission to group $GroupName on $ListName"
    Set-PnPListPermission `
        -Identity $ListName `
        -Group $GroupName `
        -AddRole "Edit"
}

Write-Host "`nAll permissions applied successfully." -ForegroundColor Green