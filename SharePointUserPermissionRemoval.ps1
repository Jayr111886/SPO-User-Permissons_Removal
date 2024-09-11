# Define the user to remove
$userEmail = "user@domain.com"  # Set the user email whose permissions will be removed
$libraryName = "Shared Documents"  # The library where permissions will be removed
$reportOnly = $true  # Set to $true for report only with NO changes; set to $false to commit and make changes
$outputCSV = "./Permission_Removal_Report.csv"  # Path to store the report of actions
$errorCSV = "./Permission_Failure_Report.csv"  # Path to store the report of errors

# Connect to the SharePoint Online Admin site to get all sites
$adminSiteUrl = "https://yourtenant-admin.sharepoint.com"
Connect-PnPOnline -Url $adminSiteUrl -Interactive

# Get all site collections in the tenant
$sites = Get-PnPTenantSite -Detailed

# Function to remove permissions for a user from all folders in a document library
function Remove-UserPermissions {
    param ($libraryName)

    # Get all folders in the document library recursively
    $listItems = Get-PnPListItem -List $libraryName -PageSize 1000 -Query "<View Scope='RecursiveAll'><Query><Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>1</Value></Eq></Where></Query></View>"

    foreach ($item in $listItems) {
        $itemUrl = $item.FieldValues["FileRef"]
        $itemType = $item.FileSystemObjectType

        Write-Host "Checking permissions for item: $itemUrl [$itemType]"

        # Check if the item has unique permissions
        if (Get-PnPProperty -ClientObject $item -Property "HasUniqueRoleAssignments") {
            Write-Host "Unique permissions found for: $itemUrl"

            # Report only mode
            $report = [PSCustomObject]@{
                SiteUrl  = $site.Url
                ItemUrl  = $itemUrl
                Action   = if ($reportOnly) { "ReportOnly - Would Remove" } else { "Permissions Removed" }
                Status   = if ($reportOnly) { "Unique Permissions Found" } else { "Success" }
            }
            $report | Export-Csv -Path $outputCSV -NoTypeInformation -Append

            if (-not $reportOnly) {
                try {
                    Write-Host "Removing permissions for $userEmail from item: $itemUrl"
                    Set-PnPListItemPermission -List $libraryName -Identity $item.Id -User $userEmail -RemoveRole "Contribute"
                }
                catch {
                    Write-Host "Error removing permissions: $($_.Exception.Message)" -ForegroundColor Red
                    $report.Status = $_.Exception.Message
                    $report | Export-Csv -Path $outputCSV -NoTypeInformation -Append
                }
            }
        }
    }
}

# Function to log site errors to the CSV file
function Log-SiteFailure {
    param ($siteUrl, $errorMessage)

    $errorLog = [PSCustomObject]@{
        SiteUrl = $siteUrl
        Reason  = $errorMessage
    }
    $errorLog | Export-Csv -Path $errorCSV -NoTypeInformation -Append
}

# Loop through each site and remove permissions from all folders in the 'Shared Documents' library
foreach ($site in $sites) {
    try {
        # Connect to the individual site
        Connect-PnPOnline -Url $site.Url -Interactive

        # Attempt to remove permissions from 'Shared Documents' if it exists
        Write-Host "Processing site: $($site.Url)"
        Remove-UserPermissions -libraryName $libraryName
    }
    catch {
        Write-Host "Skipping site $($site.Url) due to error: $($_.Exception.Message)" -ForegroundColor Yellow
        # Log the error to the failure CSV
        Log-SiteFailure -siteUrl $site.Url -errorMessage $_.Exception.Message
    }
    finally {
        # Disconnect from the current site
        Disconnect-PnPOnline
    }
}

# Disconnect from SharePoint Online Admin
Disconnect-PnPOnline
