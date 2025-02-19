# SharePoint Version Management and Cleanup Script
# Purpose: Manages version settings and cleanup for both tenant and existing sites

param (
    [Parameter(Mandatory = $false)]
    [string]$TenantName
)

function Initialize-SPOConnection {
    param (
        [string]$TenantName
    )
    
    # Check PowerShell version and warn if using PS 7
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "Notice: You're running PowerShell $($PSVersionTable.PSVersion). Some SharePoint cmdlets may work better in Windows PowerShell 5.1" -ForegroundColor Yellow
    }
    
    $module = Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell"
    if (-not $module) {
        Write-Host "Installing SharePoint Online Management Shell..." -ForegroundColor Yellow
        Install-Module -Name "Microsoft.Online.SharePoint.PowerShell" -Force -AllowClobber -Scope CurrentUser
    }
    
    Import-Module "Microsoft.Online.SharePoint.PowerShell" -UseWindowsPowerShell -DisableNameChecking
    
    if (-not $TenantName) {
        $TenantName = Read-Host "Enter your SharePoint tenant name (e.g., 'company' without the URL)"
    }
    
    if ($TenantName -match "https://([^-]+)-admin\.sharepoint\.com") {
        $TenantName = $matches[1]
    }
    elseif ($TenantName -match "https://([^.]+)\.sharepoint\.com") {
        $TenantName = $matches[1]
    }
    
    $adminSiteUrl = "https://$TenantName-admin.sharepoint.com"
    Write-Host "Connecting to: $adminSiteUrl" -ForegroundColor Cyan
    
    try {
        # Try to use modern authentication
        Connect-SPOService -Url $adminSiteUrl -ModernAuth $true -ErrorAction Stop
        return $adminSiteUrl
    }
    catch {
        Write-Host "Modern authentication failed, trying interactive login..." -ForegroundColor Yellow
        try {
            Connect-SPOService -Url $adminSiteUrl -ErrorAction Stop
            return $adminSiteUrl
        }
        catch {
            Write-Host "Connection failed. Please check:" -ForegroundColor Red
            Write-Host "1. Your tenant name is correct" -ForegroundColor Yellow
            Write-Host "2. You have admin permissions" -ForegroundColor Yellow
            Write-Host "3. You can reach $adminSiteUrl" -ForegroundColor Yellow
            Write-Host "`nError details: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
}

function Enable-TenantIntelligentVersioning {
    try {
        # Get current tenant settings
        Write-Host "`nChecking current Intelligent Versioning status..." -ForegroundColor Cyan
        $tenantSettings = Get-SPOTenant
        
        if ($tenantSettings.EnableAutoExpirationVersionTrim) {
            Write-Host "Intelligent Versioning is already enabled at tenant level" -ForegroundColor Yellow
            return
        }
        
        Write-Host "Enabling Intelligent Versioning at tenant level..." -ForegroundColor Cyan
        Set-SPOTenant -EnableAutoExpirationVersionTrim $true
        Write-Host "Successfully enabled Intelligent Versioning at tenant level" -ForegroundColor Green
    }
    catch {
        Write-Host "Error accessing/enabling tenant Intelligent Versioning: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Enable-ExistingSitesIntelligentVersioning {
    param (
        [array]$SelectedSites
    )
    
    $successCount = 0
    $failCount = 0
    $alreadyEnabledCount = 0
    
    foreach ($site in $SelectedSites) {
        Write-Host "`nChecking site: $($site.Url)" -ForegroundColor Cyan
        
        try {
            # Get current site settings
            $siteSettings = Get-SPOSite -Identity $site.Url
            
            if ($siteSettings.EnableAutoExpirationVersionTrim) {
                Write-Host "Intelligent Versioning is already enabled for this site" -ForegroundColor Yellow
                $alreadyEnabledCount++
                continue
            }
            
            Write-Host "Enabling Intelligent Versioning..." -ForegroundColor Cyan
            Set-SPOSite -Identity $site.Url -EnableAutoExpirationVersionTrim $true -Confirm:$false
            Write-Host "Successfully enabled Intelligent Versioning" -ForegroundColor Green
            $successCount++
        }
        catch {
            Write-Host "Error accessing/enabling Intelligent Versioning: $($_.Exception.Message)" -ForegroundColor Red
            $failCount++
        }
    }
    
    Write-Host "`nSummary:" -ForegroundColor Cyan
    if ($alreadyEnabledCount -gt 0) {
        Write-Host "Intelligent Versioning already enabled on $alreadyEnabledCount sites" -ForegroundColor Yellow
    }
    if ($successCount -gt 0) {
        Write-Host "Successfully enabled Intelligent Versioning on $successCount sites" -ForegroundColor Green
    }
    if ($failCount -gt 0) {
        Write-Host "Failed to enable Intelligent Versioning on $failCount sites" -ForegroundColor Red
    }
}
function Get-SiteSelection {
    try {
        $siteCollections = Get-SPOSite -Limit All -ErrorAction Stop
        if (-not $siteCollections) {
            Write-Host "No SharePoint sites found in the tenant." -ForegroundColor Red
            return $null
        }

        $sites = @()
        $index = 1

        Write-Host "`nAvailable SharePoint Sites:" -ForegroundColor Cyan
        $siteCollections | ForEach-Object {
            Write-Host "$index. $($_.Title) - $($_.Url)" -ForegroundColor White
            $sites += [PSCustomObject]@{
                Index           = $index
                Url             = $_.Url
                Title           = $_.Title
                # Add normalized versions of title and URL for matching
                NormalizedTitle = ($_.Title -replace '[^a-zA-Z0-9]', '').ToLower()
                NormalizedUrl   = ($_.Url -replace '[^a-zA-Z0-9]', '').ToLower()
                RelativePath    = ($_.Url -replace 'https://[^/]+/', '').ToLower()
            }
            $index++
        }

        do {
            Write-Host "`nSelect sites by:" -ForegroundColor Yellow
            Write-Host "- Numbers (e.g., 1,2,3)" -ForegroundColor Yellow
            Write-Host "- Site names (e.g., Projects, Archive)" -ForegroundColor Yellow
            Write-Host "- Full or partial URLs" -ForegroundColor Yellow
            Write-Host "- 'all' for all sites" -ForegroundColor Yellow
            Write-Host "- 'q' to quit" -ForegroundColor Yellow
            Write-Host "`nEnter your selection:" -ForegroundColor Yellow
            $selection = Read-Host

            if ($selection -eq 'q') {
                Write-Host "Exiting script" -ForegroundColor Yellow
                return $null
            }

            $selectedSites = @()

            if ($selection -eq 'all') {
                $selectedSites = $sites
            }
            else {
                $selections = $selection -split ',' | ForEach-Object { $_.Trim() }
                $invalidSelections = @()

                foreach ($sel in $selections) {
                    $site = $null
                    $normalizedSelection = ($sel -replace '[^a-zA-Z0-9]', '').ToLower()
                    
                    # Try parsing as number first
                    if ([int]::TryParse($sel, [ref]$null)) {
                        $site = $sites | Where-Object { $_.Index -eq [int]$sel }
                    }
                    else {
                        # Try matching by normalized values
                        $matches = $sites | Where-Object { 
                            $_.NormalizedTitle -like "*$normalizedSelection*" -or
                            $_.NormalizedUrl -like "*$normalizedSelection*" -or
                            $_.RelativePath -like "*$($sel.ToLower())*" -or
                            # Add partial matching for site parts
                            ($_.RelativePath -split '/') -contains $sel.ToLower() -or
                            # Match last part of URL path
                            ($_.Url -split '/')[-1].ToLower() -eq $sel.ToLower()
                        }

                        if ($matches.Count -eq 1) {
                            $site = $matches
                        }
                        elseif ($matches.Count -gt 1) {
                            Write-Host "`nMultiple matches found for '$sel'. Please select one:" -ForegroundColor Yellow
                            $matches | ForEach-Object {
                                Write-Host "- [$($_.Index)] $($_.Title) - $($_.Url)" -ForegroundColor White
                            }
                            $subSelection = Read-Host "`nEnter the number in brackets"
                            if ([int]::TryParse($subSelection, [ref]$null)) {
                                $site = $matches | Where-Object { $_.Index -eq [int]$subSelection }
                            }
                            if (-not $site) {
                                $invalidSelections += $sel
                            }
                        }
                    }

                    if ($site) {
                        # Avoid duplicates
                        $selectedSites += $site | Where-Object { $selectedSites.Url -notcontains $_.Url }
                    }
                    else {
                        $invalidSelections += $sel
                    }
                }

                if ($invalidSelections) {
                    Write-Host "`nWarning: Could not find matches for: $($invalidSelections -join ', ')" -ForegroundColor Yellow
                    Write-Host "Try using the site number or exact site name from the list above." -ForegroundColor Yellow
                    continue
                }
            }

            if ($selectedSites.Count -eq 0) {
                Write-Host "`nNo valid sites selected. Please try again." -ForegroundColor Yellow
                continue
            }

            # Show selected sites for confirmation
            Write-Host "`nSelected Sites:" -ForegroundColor Cyan
            $selectedSites | ForEach-Object {
                Write-Host "- $($_.Title) - $($_.Url)" -ForegroundColor White
            }
            Write-Host "`nTotal sites selected: $($selectedSites.Count)" -ForegroundColor Cyan

            $confirm = Read-Host "`nProceed with these sites? (y/n)"
            if ($confirm -eq 'y') {
                return $selectedSites
            }
            else {
                Write-Host "`nPlease select sites again." -ForegroundColor Yellow
            }
        } while ($true)
    }
    catch {
        Write-Host "Error retrieving SharePoint sites: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Start-VersionCleanup {
    param (
        [string]$DeletionMode,
        [int]$DaysToKeep,
        [int]$VersionsToKeep,
        [array]$SelectedSites
    )
    
    foreach ($site in $SelectedSites) {
        Write-Host "`nProcessing site: $($site.Url)" -ForegroundColor Cyan
        
        try {
            if ($DeletionMode -eq "days") {
                New-SPOSiteFileVersionBatchDeleteJob -Identity $site.Url -DeleteBeforeDays $DaysToKeep -Confirm:$false
                Write-Host "Created batch delete job for versions older than $DaysToKeep days" -ForegroundColor Green
            }
            elseif ($DeletionMode -eq "versions") {
                New-SPOSiteFileVersionBatchDeleteJob -Identity $site.Url -MajorVersionLimit $VersionsToKeep -Confirm:$false
                Write-Host "Created batch delete job keeping last $VersionsToKeep versions" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "Error processing site $($site.Url): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function Get-IntelligentVersioningStatus {
    try {
        Write-Host "`nChecking Intelligent Versioning Status..." -ForegroundColor Cyan
        
        # Check tenant-level setting
        $tenantSettings = Get-SPOTenant
        Write-Host "`nTenant-Level Status:" -ForegroundColor Yellow
        Write-Host "Intelligent Versioning is $(if ($tenantSettings.EnableAutoExpirationVersionTrim) {'enabled'} else {'disabled'})" -ForegroundColor $(if ($tenantSettings.EnableAutoExpirationVersionTrim) { 'Green' } else { 'Red' })
        
        # Get site collections and their status
        $siteCollections = Get-SPOSite -Limit All -ErrorAction Stop
        
        Write-Host "`nSite-Level Status:" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Yellow
        
        $enabledCount = 0
        $disabledCount = 0
        
        foreach ($site in $siteCollections) {
            $status = if ($site.EnableAutoExpirationVersionTrim) { 'enabled' } else { 'disabled' }
            $color = if ($site.EnableAutoExpirationVersionTrim) { 'Green' } else { 'Red' }
            Write-Host "$($site.Url): $status" -ForegroundColor $color
            
            if ($site.EnableAutoExpirationVersionTrim) {
                $enabledCount++
            }
            else {
                $disabledCount++
            }
        }
        
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total sites: $($siteCollections.Count)" -ForegroundColor White
        Write-Host "Sites with Intelligent Versioning enabled: $enabledCount" -ForegroundColor Green
        Write-Host "Sites with Intelligent Versioning disabled: $disabledCount" -ForegroundColor $(if ($disabledCount -gt 0) { 'Red' } else { 'Green' })
    }
    catch {
        Write-Host "Error checking Intelligent Versioning status: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-PreservationHoldLibraryStatus { 
    param (
        [array]$SelectedSites
    )
    
    try {
        # Prompt for Client ID
        $ClientId = Read-Host "Enter Client ID for connection"
        
        Write-Host "`nChecking Preservation Hold Library Status..." -ForegroundColor Cyan
        $selectedSites = Get-SiteSelection
        foreach ($site in $SelectedSites) {
            Write-Host "`nSite: $($site.Url)" -ForegroundColor Yellow
            
            try {
                # Connect to the site using the registered app
                Connect-PnPOnline -Url $site.Url -Interactive -ClientId $ClientId
                
                # Check if the Preservation Hold Library exists
                $library = Get-PnPList -Identity "Preservation Hold Library" -ErrorAction SilentlyContinue
                
                if ($null -eq $library) {
                    Write-Host "Preservation Hold Library not found" -ForegroundColor Yellow
                    continue
                }
                
                # Get library stats
                $items = Get-PnPListItem -List "Preservation Hold Library" -PageSize 2000
                
                # Calculate total size using the correct field name with progress bar
                $totalSize = 0
                $processedItems = 0
                $totalItems = $items.Count

                $items | ForEach-Object {
                    $processedItems++
                    $percentComplete = [math]::Round(($processedItems / $totalItems) * 100, 2)
                    
                    Write-Progress -Activity "Calculating Library Size" `
                        -Status "Processing item $processedItems of $totalItems" `
                        -PercentComplete $percentComplete `
                        -CurrentOperation "Current Size: $([math]::Round($totalSize/1MB, 2)) MB"

                    if ($_.FieldValues.FileLeafRef) {
                        # Only process actual files
                        $totalSize += [long]$_.FieldValues.File_x0020_Size
                    }
                }
                
                # Clear the progress bar
                Write-Progress -Activity "Calculating Library Size" -Completed
                
                $sizeInMB = [math]::Round($totalSize / 1MB, 2)
                
                Write-Host "Total Items: $($items.Count)" -ForegroundColor Green
                Write-Host "Total Size: $sizeInMB MB" -ForegroundColor Green
                
                # Disconnect from the site
                Disconnect-PnPOnline
            }
            catch {
                Write-Host "Error checking site $($site.Url): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "Error checking Preservation Hold Library status: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Get-PreservationHoldLibraryStatus-App { 
    param (
        [array]$SelectedSites,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$Tenant,
        [string]$ReportPath = ".\PreservationHoldLibraryReport.csv"
    )
    
    try {
        # Prompt for ClientId if not provided
        if ([string]::IsNullOrEmpty($ClientId)) {
            $ClientId = Read-Host -Prompt "Enter the Client ID"
        }

        # Prompt for Thumbprint if not provided
        if ([string]::IsNullOrEmpty($Thumbprint)) {
            $Thumbprint = Read-Host -Prompt "Enter the Certificate Thumbprint"
        }

        # Prompt for Tenant if not provided
        if ([string]::IsNullOrEmpty($Tenant)) {
            $Tenant = Read-Host -Prompt "Enter the Tenant name (e.g., contoso.onmicrosoft.com)"
        }

        # Prompt for Sites if not provided
        if ($null -eq $SelectedSites -or $SelectedSites.Count -eq 0) {
            $SelectedSites = Get-SiteSelection
        }

        Write-Host "`nChecking Preservation Hold Library Status..." -ForegroundColor Cyan
        
        # Create array to store results
        $results = @()
        
        foreach ($site in $SelectedSites) {
            Write-Host "`nSite: $($site.Url)" -ForegroundColor Yellow
            
            try {
                # Connect to each site using certificate authentication with tenant
                Connect-PnPOnline -Url $site.Url -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
                
                $siteResult = [PSCustomObject]@{
                    SiteUrl       = $site.Url
                    Title         = $site.Title
                    LibraryExists = $false
                    ItemCount     = 0
                    SizeInMB      = 0
                    LastChecked   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Status        = "Success"
                    ErrorMessage  = ""
                }
                
                # Check if the Preservation Hold Library exists
                $library = Get-PnPList -Identity "Preservation Hold Library" -ErrorAction SilentlyContinue
                
                if ($null -eq $library) {
                    Write-Host "Preservation Hold Library not found" -ForegroundColor Yellow
                    $siteResult.Status = "Library Not Found"
                }
                else {
                    $siteResult.LibraryExists = $true
                    
                    # Get library stats
                    $items = Get-PnPListItem -List "Preservation Hold Library" -PageSize 2000
                    
                    # Calculate total size using the correct field name with progress bar
                    $totalSize = 0
                    $processedItems = 0
                    $totalItems = $items.Count

                    $items | ForEach-Object {
                        $processedItems++
                        $percentComplete = [math]::Round(($processedItems / $totalItems) * 100, 2)
                        
                        Write-Progress -Activity "Calculating Library Size" `
                            -Status "Processing item $processedItems of $totalItems" `
                            -PercentComplete $percentComplete `
                            -CurrentOperation "Current Size: $([math]::Round($totalSize/1MB, 2)) MB"

                        if ($_.FieldValues.FileLeafRef) {
                            # Only process actual files
                            $totalSize += [long]$_.FieldValues.File_x0020_Size
                        }
                    }
                    
                    # Clear the progress bar
                    Write-Progress -Activity "Calculating Library Size" -Completed
                    
                    $sizeInMB = [math]::Round($totalSize / 1MB, 2)
                    
                    $siteResult.ItemCount = $items.Count
                    $siteResult.SizeInMB = $sizeInMB
                    
                    Write-Host "Total Items: $($items.Count)" -ForegroundColor Green
                    Write-Host "Total Size: $sizeInMB MB" -ForegroundColor Green
                }
                
                # Add result to array
                $results += $siteResult
                
                # Disconnect before moving to next site
                Disconnect-PnPOnline
            }
            catch {
                Write-Host "Error checking site $($site.Url): $($_.Exception.Message)" -ForegroundColor Red
                
                # Add error result to array
                $results += [PSCustomObject]@{
                    SiteUrl       = $site.Url
                    Title         = $site.Title
                    LibraryExists = $false
                    ItemCount     = 0
                    SizeInMB      = 0
                    LastChecked   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    Status        = "Error"
                    ErrorMessage  = $_.Exception.Message
                }
            }
        }
        
        # Export results to CSV
        $results | Export-Csv -Path $ReportPath -NoTypeInformation
        
        # Display summary
        Write-Host "`nSummary:" -ForegroundColor Cyan
        Write-Host "Total Sites Checked: $($results.Count)" -ForegroundColor White
        Write-Host "Sites with Library: $($results.Where({$_.LibraryExists}).Count)" -ForegroundColor Green
        Write-Host "Sites with Errors: $($results.Where({$_.Status -eq 'Error'}).Count)" -ForegroundColor Red
        Write-Host "Total Items Across All Libraries: $($results | Measure-Object -Property ItemCount -Sum | Select-Object -ExpandProperty Sum)" -ForegroundColor White
        Write-Host "Total Size Across All Libraries: $([math]::Round(($results | Measure-Object -Property SizeInMB -Sum | Select-Object -ExpandProperty Sum), 2)) MB" -ForegroundColor White
        Write-Host "`nReport exported to: $ReportPath" -ForegroundColor Green
        
    }
    catch {
        Write-Host "Error checking Preservation Hold Library status: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function DownloadPreservationHoldLibrary {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        [string]$PreservationHoldLibraryName = "Preservation Hold Library",
        [string]$DownloadPath = ".\PreservationHoldFiles",
        [int]$PageSize = 2000
    )
    
    function Get-BlobMetadata {
        param (
            [string]$BlobUrl,
            [string]$SasToken
        )
        try {
            $headers = @{ "x-ms-version" = "2019-12-12" }
            $response = Invoke-RestMethod -Uri "$BlobUrl$SasToken" -Method Head -Headers $headers
            return $response.Headers | Where-Object { $_.Key.StartsWith('x-ms-meta-') }
        }
        catch {
            Write-Host "Error retrieving blob metadata: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    }

    function Clean-FileName {
        param (
            [string]$FileName
        )
        
        # Pattern to match: anything followed by underscore and GUID-like string and date
        if ($FileName -match '^(.+?)_[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}\d{4}-\d{2}-\d{2}T') {
            return $matches[1] + [System.IO.Path]::GetExtension($FileName)
        }
        
        # Alternative pattern for filenames that start with dates
        if ($FileName -match '^(\d{8}_\d{6})_[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}') {
            return $matches[1] + [System.IO.Path]::GetExtension($FileName)
        }
        
        # If no patterns match, return original filename
        return $FileName
    }

    function Test-AzureBlob {
        param (
            [string]$BlobUrl,
            [string]$SasToken
        )
        try {
            $headers = @{ "x-ms-version" = "2019-12-12" }
            Invoke-RestMethod -Uri "$BlobUrl$SasToken" -Method Head -Headers $headers -ErrorAction Stop
            return $true
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                return $false
            }
            throw
        }
    }

    function Get-UniqueFileName {
        param (
            [string]$FilePath,
            [string]$FileName
        )
        
        $fileNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $extension = [System.IO.Path]::GetExtension($FileName)
        $counter = 1
        $newPath = Join-Path $FilePath $FileName
        
        while (Test-Path $newPath) {
            $newFileName = "${fileNameOnly}_${counter}${extension}"
            $newPath = Join-Path $FilePath $newFileName
            $counter++
        }
        
        return (Split-Path $newPath -Leaf)
    }
    
    try {
        # Prompt for Client ID
        $ClientId = Read-Host "Enter Client ID for connection"
        
        # Prompt for duplicate handling preference
        Write-Host "`nHow would you like to handle duplicate files?" -ForegroundColor Cyan
        Write-Host "1: Skip existing files" -ForegroundColor Yellow
        Write-Host "2: Overwrite existing files" -ForegroundColor Yellow
        Write-Host "3: Rename files to avoid conflicts" -ForegroundColor Yellow
        
        do {
            $duplicateChoice = Read-Host "`nEnter your choice (1-3)"
        } while ($duplicateChoice -notmatch '^[1-3]$')
        
        $DuplicateHandling = switch ($duplicateChoice) {
            "1" { "Skip" }
            "2" { "Overwrite" }
            "3" { "Rename" }
        }
        
        Write-Host "Selected duplicate handling method: $DuplicateHandling" -ForegroundColor Green
        
        # Prompt for Azure Blob Storage options
        $uploadToAzure = (Read-Host "Would you like to upload files to Azure Blob Storage? (Y/N)") -eq 'Y'
        
        if ($uploadToAzure) {
            # Get Azure storage details
            $storageAccountName = Read-Host "Enter the Azure Storage Account name"
            $containerName = Read-Host "Enter the Container name"
            $sasToken = Read-Host "Enter the SAS token"
            
            # Ensure SAS token starts with '?'
            if (-not $sasToken.StartsWith('?')) {
                $sasToken = "?" + $sasToken
            }
            
            # Construct the blob storage endpoint
            $blobEndpoint = "https://$storageAccountName.blob.core.windows.net"
            $containerUrl = "$blobEndpoint/$containerName"
            
            # Test Azure connection
            try {
                $testUrl = "$containerUrl/test.txt$sasToken"
                $headers = @{
                    'x-ms-blob-type' = 'BlockBlob'
                    'Content-Type'   = 'text/plain'
                }
                $testContent = [System.Text.Encoding]::UTF8.GetBytes("test")
                
                Invoke-RestMethod -Uri $testUrl -Method Put -Headers $headers -Body $testContent
                Write-Host "Successfully verified Azure Blob Storage access" -ForegroundColor Green
                
                # Clean up test file
                Invoke-RestMethod -Uri $testUrl -Method Delete
            }
            catch {
                Write-Host "Error accessing Azure Blob Storage: $($_.Exception.Message)" -ForegroundColor Red
                $continue = Read-Host "Would you like to continue without Azure upload? (Y/N)"
                if ($continue -ne 'Y') {
                    return
                }
                $uploadToAzure = $false
            }
        }
        
        # Create download directory if it doesn't exist
        $DownloadPath = (Resolve-Path $DownloadPath -ErrorAction SilentlyContinue).Path
        if (-not $DownloadPath) {
            $DownloadPath = (New-Item -ItemType Directory -Path $DownloadPath).FullName
            Write-Host "Created download directory at $DownloadPath" -ForegroundColor Green
        }
        
        # Connect to SharePoint
        Write-Host "Connecting to site: $SiteUrl" -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
        
        # Check if the Preservation Hold Library exists
        $library = Get-PnPList -Identity $PreservationHoldLibraryName -ErrorAction SilentlyContinue
        if ($null -eq $library) {
            Write-Host "Preservation Hold Library not found on this site" -ForegroundColor Yellow
            return
        }
        
        # Get all items
        Write-Host "Retrieving items from Preservation Hold Library..." -ForegroundColor Cyan
        $items = Get-PnPListItem -List $PreservationHoldLibraryName -PageSize $PageSize
        
        if ($items.Count -eq 0) {
            Write-Host "No items found in the Preservation Hold Library" -ForegroundColor Yellow
            return
        }
        
        Write-Host "Found $($items.Count) items in the Preservation Hold Library" -ForegroundColor Cyan
        
        # Initialize counters
        $SuccessCount = 0
        $ErrorCount = 0
        $ProcessedCount = 0
        $SkippedCount = 0
        $UploadSuccessCount = 0
        $UploadErrorCount = 0
        $failedItems = @()
        $failedUploads = @()
        
        # Calculate total size
        $totalBytes = 0
        $downloadedBytes = 0
        $uploadedBytes = 0
        
        $items | ForEach-Object {
            if ($_.FieldValues.FileLeafRef) {
                $totalBytes += [long]$_.FieldValues.File_x0020_Size
            }
        }
        
        Write-Host "Total size to process: $([math]::Round($totalBytes/1MB, 2)) MB" -ForegroundColor Cyan
        
        # Process items
        foreach ($item in $items) {
            $ProcessedCount++
            
            if (!$item.FieldValues.FileLeafRef) {
                $SkippedCount++
                continue
            }
            
            try {
                $fileName = Clean-FileName $item.FieldValues.FileLeafRef
                $fileSize = [long]$item.FieldValues.File_x0020_Size
                $filePath = Join-Path $DownloadPath $fileName
                
                # Handle local file duplicates
                if (Test-Path $filePath) {
                    switch ($DuplicateHandling) {
                        "Skip" {
                            Write-Host "Skipping existing file: $fileName" -ForegroundColor Yellow
                            $SkippedCount++
                            continue
                        }
                        "Rename" {
                            $fileName = Get-UniqueFileName -FilePath $DownloadPath -FileName $fileName
                            $filePath = Join-Path $DownloadPath $fileName
                            Write-Host "Renaming to avoid conflict: $fileName" -ForegroundColor Yellow
                        }
                        "Overwrite" {
                            Write-Host "Overwriting existing file: $fileName" -ForegroundColor Yellow
                            Remove-Item $filePath -Force
                        }
                    }
                }
                
                # Show download progress
                $percentComplete = [math]::Min(100, [math]::Round(($ProcessedCount / $items.Count) * 100, 2))
                $remainingFiles = $items.Count - $ProcessedCount
                Write-Progress -Id 1 -Activity "Downloading Files" `
                    -Status "Downloading: $ProcessedCount of $($items.Count) ($remainingFiles remaining)" `
                    -PercentComplete $percentComplete
                
                # Download file
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $DownloadPath -Filename $fileName -AsFile -Force
                $downloadedBytes += $fileSize
                
                # Get the full path of the downloaded file
                $downloadedFilePath = Join-Path $DownloadPath $fileName
                
                # Set the file creation and modification dates to match SharePoint
                $created = $item.FieldValues.Created
                $modified = $item.FieldValues.Modified
                
                if ($created -and $modified) {
                    try {
                        $file = Get-Item $downloadedFilePath
                        $file.CreationTime = $created
                        $file.LastWriteTime = $modified
                        $file.LastAccessTime = $modified
                    }
                    catch {
                        Write-Host "Warning: Could not set dates for $fileName : $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                
                $SuccessCount++
                
                # Upload to Azure if enabled
                if ($uploadToAzure) {
                    try {
                        $blobUrl = "$containerUrl/$fileName"
                        
                        # Check if blob already exists
                        $blobExists = Test-AzureBlob -BlobUrl $blobUrl -SasToken $sasToken
                        
                        if ($blobExists) {
                            switch ($DuplicateHandling) {
                                "Skip" {
                                    Write-Host "Skipping existing blob: $fileName" -ForegroundColor Yellow
                                    continue
                                }
                                "Rename" {
                                    $azureFileName = Get-UniqueFileName -FilePath $DownloadPath -FileName $fileName
                                    $blobUrl = "$containerUrl/$azureFileName"
                                    Write-Host "Renaming blob to avoid conflict: $azureFileName" -ForegroundColor Yellow
                                }
                                "Overwrite" {
                                    Write-Host "Overwriting existing blob: $fileName" -ForegroundColor Yellow
                                }
                            }
                        }
                        
                        Write-Progress -Id 2 -Activity "Uploading to Azure" `
                            -Status "Uploading: $UploadSuccessCount of $SuccessCount" `
                            -PercentComplete ([math]::Min(100, [math]::Round(($UploadSuccessCount / $SuccessCount) * 100, 2)))
                        
                        # Extract SharePoint metadata
                        $created = $item.FieldValues.Created
                        $modified = $item.FieldValues.Modified
                        $author = $item.FieldValues.Author.Email
                        $editor = $item.FieldValues.Editor.Email

                        # Format dates to a valid format and encode properly
                        # Use a format that's safe for metadata
                        $createdDate = $created.ToString("yyyy-MM-dd HH:mm:ss")
                        $modifiedDate = $modified.ToString("yyyy-MM-dd HH:mm:ss")
                        
                        # Sanitize email addresses (remove special characters)
                        $author = $author -replace '[^\w@.-]', ''
                        $editor = $editor -replace '[^\w@.-]', ''

                        # Prepare headers with metadata
                        $headers = @{
                            'x-ms-blob-type'          = 'BlockBlob'
                            'Content-Type'            = 'application/octet-stream'
                            'x-ms-blob-cache-control' = 'no-cache'
                            # Use sanitized metadata values
                            'x-ms-meta-creationdate'  = $createdDate
                            'x-ms-meta-modifieddate'  = $modifiedDate
                            'x-ms-meta-author'        = $author
                            'x-ms-meta-editor'        = $editor
                        }
                        
                        $actualFilePath = Get-ChildItem -Path $DownloadPath -Recurse -Filter $fileName | 
                        Select-Object -ExpandProperty FullName -First 1
                        
                        if (-not $actualFilePath) {
                            throw "Could not find downloaded file: $fileName"
                        }
                        
                        $fileContent = [System.IO.File]::ReadAllBytes($actualFilePath)
                        $uploadResponse = Invoke-RestMethod -Uri "$blobUrl$sasToken" -Method Put -Headers $headers -Body $fileContent
                        
                        $uploadedBytes += $fileSize
                        $UploadSuccessCount++
                    }
                    catch {
                        $UploadErrorCount++
                        $failedUploads += $fileName
                        Write-Host "Failed to upload $fileName : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            catch {
                $ErrorCount++
                $failedItems += $fileName
                Write-Host "Failed to download: $fileName" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        Write-Progress -Id 1 -Activity "Downloading Files" -Completed
        if ($uploadToAzure) {
            Write-Progress -Id 2 -Activity "Uploading to Azure" -Completed
        }
        
        # Display summary
        Write-Host "`nOperation Summary:" -ForegroundColor Cyan
        Write-Host "Total Items: $($items.Count)" -ForegroundColor White
        Write-Host "Successfully Downloaded: $SuccessCount" -ForegroundColor Green
        Write-Host "Download Size: $([math]::Round($downloadedBytes/1MB, 2)) MB" -ForegroundColor Cyan
        
        if ($uploadToAzure) {
            Write-Host "`nAzure Upload Results:" -ForegroundColor Cyan
            Write-Host "Successfully Uploaded: $UploadSuccessCount" -ForegroundColor Green
            Write-Host "Upload Size: $([math]::Round($uploadedBytes/1MB, 2)) MB" -ForegroundColor Cyan
            
            if ($UploadErrorCount -gt 0) {
                Write-Host "`nFailed Uploads ($UploadErrorCount):" -ForegroundColor Red
                foreach ($failedUpload in $failedUploads) {
                    Write-Host "- $failedUpload" -ForegroundColor Red
                }
            }
        }
        
        if ($ErrorCount -gt 0) {
            Write-Host "`nFailed Downloads ($ErrorCount):" -ForegroundColor Red
            foreach ($failedItem in $failedItems) {
                Write-Host "- $failedItem" -ForegroundColor Red
            }
        }
        
        if ($SkippedCount -gt 0) {
            Write-Host "Skipped Items: $SkippedCount" -ForegroundColor Yellow
        }
        
        # Create log file
        $logFile = Join-Path $DownloadPath "operation_log.txt"
        $logContent = @"
Operation Summary for $SiteUrl
Date: $(Get-Date)
Total Items: $($items.Count)
Successfully Downloaded: $SuccessCount
Failed Downloads: $ErrorCount
Skipped Items: $SkippedCount
Download Size: $([math]::Round($downloadedBytes/1MB, 2)) MB

Azure Upload Status: $(if ($uploadToAzure) { "Enabled" } else { "Disabled" })
$(if ($uploadToAzure) {
"Successfully Uploaded: $UploadSuccessCount
Failed Uploads: $UploadErrorCount
Upload Size: $([math]::Round($uploadedBytes/1MB, 2)) MB

Failed Uploads:
$($failedUploads -join "`n")"
})

Failed Downloads:
$($failedItems -join "`n")
"@
        $logContent | Out-File $logFile
        
        Write-Host "`nOperation log saved to: $logFile" -ForegroundColor Green
    }
    catch {
        Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    finally {
        # Disconnect PnP connection
        try {
            Disconnect-PnPOnline
            Write-Host "Disconnected from SharePoint" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error disconnecting from SharePoint: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

function DownloadPreservationHoldLibrary-App {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        [string]$PreservationHoldLibraryName = "Preservation Hold Library",
        [string]$DownloadPath = ".\PreservationHoldFiles",
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$Tenant,
        [int]$PageSize = 2000
    )
    
    function Get-BlobMetadata {
        param (
            [string]$BlobUrl,
            [string]$SasToken
        )
        try {
            $headers = @{ "x-ms-version" = "2019-12-12" }
            $response = Invoke-RestMethod -Uri "$BlobUrl$SasToken" -Method Head -Headers $headers
            return $response.Headers | Where-Object { $_.Key.StartsWith('x-ms-meta-') }
        }
        catch {
            Write-Host "Error retrieving blob metadata: $($_.Exception.Message)" -ForegroundColor Red
            return $null
        }
    }

    function Clean-FileName {
        param (
            [string]$FileName
        )
        
        # Pattern to match: anything followed by underscore and GUID-like string and date
        if ($FileName -match '^(.+?)_[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}\d{4}-\d{2}-\d{2}T') {
            return $matches[1] + [System.IO.Path]::GetExtension($FileName)
        }
        
        # Alternative pattern for filenames that start with dates
        if ($FileName -match '^(\d{8}_\d{6})_[A-F0-9]{8}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{4}-[A-F0-9]{12}') {
            return $matches[1] + [System.IO.Path]::GetExtension($FileName)
        }
        
        # If no patterns match, return original filename
        return $FileName
    }

    function Test-AzureBlob {
        param (
            [string]$BlobUrl,
            [string]$SasToken
        )
        try {
            $headers = @{ "x-ms-version" = "2019-12-12" }
            Invoke-RestMethod -Uri "$BlobUrl$SasToken" -Method Head -Headers $headers -ErrorAction Stop
            return $true
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 404) {
                return $false
            }
            throw
        }
    }

    function Get-UniqueFileName {
        param (
            [string]$FilePath,
            [string]$FileName
        )
        
        $fileNameOnly = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
        $extension = [System.IO.Path]::GetExtension($FileName)
        $counter = 1
        $newPath = Join-Path $FilePath $FileName
        
        while (Test-Path $newPath) {
            $newFileName = "${fileNameOnly}_${counter}${extension}"
            $newPath = Join-Path $FilePath $newFileName
            $counter++
        }
        
        return (Split-Path $newPath -Leaf)
    }
    
    try {
        # Prompt for ClientId if not provided
        if ([string]::IsNullOrEmpty($ClientId)) {
            $ClientId = Read-Host -Prompt "Enter the Client ID"
        }

        # Prompt for Thumbprint if not provided
        if ([string]::IsNullOrEmpty($Thumbprint)) {
            $Thumbprint = Read-Host -Prompt "Enter the Certificate Thumbprint"
        }

        # Prompt for Tenant if not provided
        if ([string]::IsNullOrEmpty($Tenant)) {
            $Tenant = Read-Host -Prompt "Enter the Tenant name (e.g., contoso.onmicrosoft.com)"
        }
        
        # Prompt for duplicate handling preference
        Write-Host "`nHow would you like to handle duplicate files?" -ForegroundColor Cyan
        Write-Host "1: Skip existing files" -ForegroundColor Yellow
        Write-Host "2: Overwrite existing files" -ForegroundColor Yellow
        Write-Host "3: Rename files to avoid conflicts" -ForegroundColor Yellow
        
        do {
            $duplicateChoice = Read-Host "`nEnter your choice (1-3)"
        } while ($duplicateChoice -notmatch '^[1-3]$')
        
        $DuplicateHandling = switch ($duplicateChoice) {
            "1" { "Skip" }
            "2" { "Overwrite" }
            "3" { "Rename" }
        }
        
        Write-Host "Selected duplicate handling method: $DuplicateHandling" -ForegroundColor Green
        
        # Prompt for Azure Blob Storage options
        $uploadToAzure = (Read-Host "Would you like to upload files to Azure Blob Storage? (Y/N)") -eq 'Y'
        
        if ($uploadToAzure) {
            # Get Azure storage details
            $storageAccountName = Read-Host "Enter the Azure Storage Account name"
            $containerName = Read-Host "Enter the Container name"
            $sasToken = Read-Host "Enter the SAS token"
            
            # Ensure SAS token starts with '?'
            if (-not $sasToken.StartsWith('?')) {
                $sasToken = "?" + $sasToken
            }
            
            # Construct the blob storage endpoint
            $blobEndpoint = "https://$storageAccountName.blob.core.windows.net"
            $containerUrl = "$blobEndpoint/$containerName"
            
            # Test Azure connection
            try {
                $testUrl = "$containerUrl/test.txt$sasToken"
                $headers = @{
                    'x-ms-blob-type' = 'BlockBlob'
                    'Content-Type'   = 'text/plain'
                }
                $testContent = [System.Text.Encoding]::UTF8.GetBytes("test")
                
                Invoke-RestMethod -Uri $testUrl -Method Put -Headers $headers -Body $testContent
                Write-Host "Successfully verified Azure Blob Storage access" -ForegroundColor Green
                
                # Clean up test file
                Invoke-RestMethod -Uri $testUrl -Method Delete
            }
            catch {
                Write-Host "Error accessing Azure Blob Storage: $($_.Exception.Message)" -ForegroundColor Red
                $continue = Read-Host "Would you like to continue without Azure upload? (Y/N)"
                if ($continue -ne 'Y') {
                    return
                }
                $uploadToAzure = $false
            }
        }
        
        # Create download directory if it doesn't exist
        $DownloadPath = (Resolve-Path $DownloadPath -ErrorAction SilentlyContinue).Path
        if (-not $DownloadPath) {
            $DownloadPath = (New-Item -ItemType Directory -Path $DownloadPath).FullName
            Write-Host "Created download directory at $DownloadPath" -ForegroundColor Green
        }
        
        # Connect to SharePoint using certificate
        Write-Host "Connecting to site: $SiteUrl" -ForegroundColor Cyan
        Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $Tenant
        
        # Check if the Preservation Hold Library exists
        $library = Get-PnPList -Identity $PreservationHoldLibraryName -ErrorAction SilentlyContinue
        if ($null -eq $library) {
            Write-Host "Preservation Hold Library not found on this site" -ForegroundColor Yellow
            return
        }
        
        # Get all items
        Write-Host "Retrieving items from Preservation Hold Library..." -ForegroundColor Cyan
        $items = Get-PnPListItem -List $PreservationHoldLibraryName -PageSize $PageSize
        
        if ($items.Count -eq 0) {
            Write-Host "No items found in the Preservation Hold Library" -ForegroundColor Yellow
            return
        }
        
        Write-Host "Found $($items.Count) items in the Preservation Hold Library" -ForegroundColor Cyan
        
        # Initialize counters
        $SuccessCount = 0
        $ErrorCount = 0
        $ProcessedCount = 0
        $SkippedCount = 0
        $UploadSuccessCount = 0
        $UploadErrorCount = 0
        $failedItems = @()
        $failedUploads = @()
        
        # Calculate total size
        $totalBytes = 0
        $downloadedBytes = 0
        $uploadedBytes = 0
        
        $items | ForEach-Object {
            if ($_.FieldValues.FileLeafRef) {
                $totalBytes += [long]$_.FieldValues.File_x0020_Size
            }
        }
        
        Write-Host "Total size to process: $([math]::Round($totalBytes/1MB, 2)) MB" -ForegroundColor Cyan
        
        # Process items
        foreach ($item in $items) {
            $ProcessedCount++
            
            if (!$item.FieldValues.FileLeafRef) {
                $SkippedCount++
                continue
            }
            
            try {
                $fileName = Clean-FileName $item.FieldValues.FileLeafRef
                $fileSize = [long]$item.FieldValues.File_x0020_Size
                $filePath = Join-Path $DownloadPath $fileName
                
                # Handle local file duplicates
                if (Test-Path $filePath) {
                    switch ($DuplicateHandling) {
                        "Skip" {
                            Write-Host "Skipping existing file: $fileName" -ForegroundColor Yellow
                            $SkippedCount++
                            continue
                        }
                        "Rename" {
                            $fileName = Get-UniqueFileName -FilePath $DownloadPath -FileName $fileName
                            $filePath = Join-Path $DownloadPath $fileName
                            Write-Host "Renaming to avoid conflict: $fileName" -ForegroundColor Yellow
                        }
                        "Overwrite" {
                            Write-Host "Overwriting existing file: $fileName" -ForegroundColor Yellow
                            Remove-Item $filePath -Force
                        }
                    }
                }
                
                # Show download progress
                $percentComplete = [math]::Min(100, [math]::Round(($ProcessedCount / $items.Count) * 100, 2))
                $remainingFiles = $items.Count - $ProcessedCount
                Write-Progress -Id 1 -Activity "Downloading Files" `
                    -Status "Downloading: $ProcessedCount of $($items.Count) ($remainingFiles remaining)" `
                    -PercentComplete $percentComplete
                
                # Download file
                Get-PnPFile -Url $item.FieldValues.FileRef -Path $DownloadPath -Filename $fileName -AsFile -Force
                $downloadedBytes += $fileSize
                
                # Get the full path of the downloaded file
                $downloadedFilePath = Join-Path $DownloadPath $fileName
                
                # Set the file creation and modification dates to match SharePoint
                $created = $item.FieldValues.Created
                $modified = $item.FieldValues.Modified
                
                if ($created -and $modified) {
                    try {
                        $file = Get-Item $downloadedFilePath
                        $file.CreationTime = $created
                        $file.LastWriteTime = $modified
                        $file.LastAccessTime = $modified
                    }
                    catch {
                        Write-Host "Warning: Could not set dates for $fileName : $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
                
                $SuccessCount++
                
                # Upload to Azure if enabled
                if ($uploadToAzure) {
                    try {
                        $blobUrl = "$containerUrl/$fileName"
                        
                        # Check if blob already exists
                        $blobExists = Test-AzureBlob -BlobUrl $blobUrl -SasToken $sasToken
                        
                        if ($blobExists) {
                            switch ($DuplicateHandling) {
                                "Skip" {
                                    Write-Host "Skipping existing blob: $fileName" -ForegroundColor Yellow
                                    continue
                                }
                                "Rename" {
                                    $azureFileName = Get-UniqueFileName -FilePath $DownloadPath -FileName $fileName
                                    $blobUrl = "$containerUrl/$azureFileName"
                                    Write-Host "Renaming blob to avoid conflict: $azureFileName" -ForegroundColor Yellow
                                }
                                "Overwrite" {
                                    Write-Host "Overwriting existing blob: $fileName" -ForegroundColor Yellow
                                }
                            }
                        }
                        
                        Write-Progress -Id 2 -Activity "Uploading to Azure" `
                            -Status "Uploading: $UploadSuccessCount of $SuccessCount" `
                            -PercentComplete ([math]::Min(100, [math]::Round(($UploadSuccessCount / $SuccessCount) * 100, 2)))
                        
                        # Extract SharePoint metadata
                        $created = $item.FieldValues.Created
                        $modified = $item.FieldValues.Modified
                        $author = $item.FieldValues.Author.Email
                        $editor = $item.FieldValues.Editor.Email

                        # Format dates to a valid format and encode properly
                        $createdDate = $created.ToString("yyyy-MM-dd HH:mm:ss")
                        $modifiedDate = $modified.ToString("yyyy-MM-dd HH:mm:ss")
                        
                        # Sanitize email addresses
                        $author = $author -replace '[^\w@.-]', ''
                        $editor = $editor -replace '[^\w@.-]', ''

                        # Prepare headers with metadata
                        $headers = @{
                            'x-ms-blob-type'          = 'BlockBlob'
                            'Content-Type'            = 'application/octet-stream'
                            'x-ms-blob-cache-control' = 'no-cache'
                            'x-ms-meta-creationdate'  = $createdDate
                            'x-ms-meta-modifieddate'  = $modifiedDate
                            'x-ms-meta-author'        = $author
                            'x-ms-meta-editor'        = $editor
                        }
                        
                        $actualFilePath = Get-ChildItem -Path $DownloadPath -Recurse -Filter $fileName | 
                        Select-Object -ExpandProperty FullName -First 1
                        
                        if (-not $actualFilePath) {
                            throw "Could not find downloaded file: $fileName"
                        }
                        
                        $fileContent = [System.IO.File]::ReadAllBytes($actualFilePath)
                        $uploadResponse = Invoke-RestMethod -Uri "$blobUrl$sasToken" -Method Put -Headers $headers -Body $fileContent
                        
                        $uploadedBytes += $fileSize
                        $UploadSuccessCount++
                    }
                    catch {
                        $UploadErrorCount++
                        $failedUploads += $fileName
                        Write-Host "Failed to upload $fileName : $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
            }
            catch {
                $ErrorCount++
                $failedItems += $fileName
                Write-Host "Failed to download: $fileName" -ForegroundColor Red
                Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        Write-Progress -Id 1 -Activity "Downloading Files" -Completed
        if ($uploadToAzure) {
            Write-Progress -Id 2 -Activity "Uploading to Azure" -Completed
        }
        
        # Display summary
        Write-Host "`nOperation Summary:" -ForegroundColor Cyan
        Write-Host "Total Items: $($items.Count)" -ForegroundColor White
        Write-Host "Successfully Downloaded: $SuccessCount" -ForegroundColor Green
        Write-Host "Download Size: $([math]::Round($downloadedBytes/1MB, 2)) MB" -ForegroundColor Cyan
        
        if ($uploadToAzure) {
            Write-Host "`nAzure Upload Results:" -ForegroundColor Cyan
            Write-Host "Successfully Uploaded: $UploadSuccessCount" -ForegroundColor Green
            Write-Host "Upload Size: $([math]::Round($uploadedBytes/1MB, 2)) MB" -ForegroundColor Cyan
            
            if ($UploadErrorCount -gt 0) {
                Write-Host "`nFailed Uploads ($UploadErrorCount):" -ForegroundColor Red
                foreach ($failedUpload in $failedUploads) {
                    Write-Host "- $failedUpload" -ForegroundColor Red
                }
            }
        }
        
        if ($ErrorCount -gt 0) {
            Write-Host "`nFailed Downloads ($ErrorCount):" -ForegroundColor Red
            foreach ($failedItem in $failedItems) {
                Write-Host "- $failedItem" -ForegroundColor Red
            }
        }
        
        if ($SkippedCount -gt 0) {
            Write-Host "Skipped Items: $SkippedCount" -ForegroundColor Yellow
        }
        
        # Create log file
        $logFile = Join-Path $DownloadPath "operation_log.txt"
        $logContent = @"
Operation Summary for $SiteUrl
Date: $(Get-Date)
Total Items: $($items.Count)
Successfully Downloaded: $SuccessCount
Failed Downloads: $ErrorCount
Skipped Items: $SkippedCount
Download Size: $([math]::Round($downloadedBytes/1MB, 2)) MB

Azure Upload Status: $(if ($uploadToAzure) { "Enabled" } else { "Disabled" })
$(if ($uploadToAzure) {
"Successfully Uploaded: $UploadSuccessCount
Failed Uploads: $UploadErrorCount
Upload Size: $([math]::Round($uploadedBytes/1MB, 2)) MB

Failed Uploads:
$($failedUploads -join "`n")"
})

Failed Downloads:
$($failedItems -join "`n")
"@
        $logContent | Out-File $logFile
        
        Write-Host "`nOperation log saved to: $logFile" -ForegroundColor Green
    }
    catch {
        Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
    finally {
        # Disconnect PnP connection
        try {
            Disconnect-PnPOnline
            Write-Host "Disconnected from SharePoint" -ForegroundColor Yellow
        }
        catch {
            Write-Host "Error disconnecting from SharePoint: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# New main script execution
try {
    # Create Logs directory if it doesn't exist
    $logPath = ".\Logs"
    if (-not (Test-Path -Path $logPath)) {
        try {
            New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        }
        catch {
            Write-Host "Error creating Logs directory: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }

    Start-Transcript -Path "$logPath\SPO_Version_Management_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $env:PNPPOWERSHELL_UPDATECHECK = "false"
    
    # Show menu with improved formatting and grouping
    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║             SharePoint Version Management Menu             ║" -ForegroundColor Yellow
    Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
    Write-Host "║  INTELLIGENT VERSIONING                                    ║" -ForegroundColor Yellow
    Write-Host "║    1. Enable for New Sites Only                            ║" -ForegroundColor Cyan
    Write-Host "║    5. Enable for All Sites                                 ║" -ForegroundColor Cyan
    Write-Host "║    6. Check Current Status                                 ║" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Yellow
    Write-Host "║  VERSION CLEANUP                                           ║" -ForegroundColor Yellow
    Write-Host "║    2. Clean Up by Age                                      ║" -ForegroundColor Cyan
    Write-Host "║    3. Clean Up by Count                                    ║" -ForegroundColor Cyan
    Write-Host "║    4. Enable Versioning + Clean Up                         ║" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Yellow
    Write-Host "║  PRESERVATION HOLD LIBRARY                                 ║" -ForegroundColor Yellow
    Write-Host "║    7. Review Contents                                      ║" -ForegroundColor Cyan
    Write-Host "║    8. Download Contents                                    ║" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Yellow
    Write-Host "║    9. Exit                                                 ║" -ForegroundColor Cyan
    Write-Host "║                                                            ║" -ForegroundColor Yellow
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    
    $choice = Read-Host "`nEnter your choice (1-9)"
    
    # Exit early if user chooses to quit
    if ($choice -eq "9") {
        Write-Host "Exiting script" -ForegroundColor Yellow
        return
    }
    
    # Validate choice before connecting
    if ($choice -notin "1", "2", "3", "4", "5", "6", "7", "8") {
        Write-Host "Invalid choice. Exiting script" -ForegroundColor Red
        return
    }
    
    # Connect to SharePoint Online only if a valid option was selected
    $adminSiteUrl = Initialize-SPOConnection -TenantName $TenantName
    
    # Process the selected option
    switch ($choice) {
        "1" { 
            Enable-TenantIntelligentVersioning
        }
        "2" { 
            $DaysToKeep = Read-Host "Enter number of days to keep versions for"
            $selectedSites = Get-SiteSelection
            if ($selectedSites) {
                Start-VersionCleanup -DeletionMode "days" -DaysToKeep $DaysToKeep -SelectedSites $selectedSites
            }
        }
        "3" { 
            $VersionsToKeep = Read-Host "Enter number of versions to keep"
            $selectedSites = Get-SiteSelection
            if ($selectedSites) {
                Start-VersionCleanup -DeletionMode "versions" -VersionsToKeep $VersionsToKeep -SelectedSites $selectedSites
            }
        }
        "4" { 
            Enable-TenantIntelligentVersioning
            $cleanupMode = Read-Host "Choose cleanup mode (days/versions)"
            if ($cleanupMode -eq "days") {
                $DaysToKeep = Read-Host "Enter number of days to keep versions for"
                $selectedSites = Get-SiteSelection
                if ($selectedSites) {
                    Start-VersionCleanup -DeletionMode "days" -DaysToKeep $DaysToKeep -SelectedSites $selectedSites
                }
            } 
            else {
                $VersionsToKeep = Read-Host "Enter number of versions to keep"
                $selectedSites = Get-SiteSelection
                if ($selectedSites) {
                    Start-VersionCleanup -DeletionMode "versions" -VersionsToKeep $VersionsToKeep -SelectedSites $selectedSites
                }
            }
        }
        "5" { 
            Enable-TenantIntelligentVersioning
            $selectedSites = Get-SiteSelection
            if ($selectedSites) {
                Enable-ExistingSitesIntelligentVersioning -SelectedSites $selectedSites
            }
        }
        "6" {
            Get-IntelligentVersioningStatus
        }
        "7" { 
            # Create submenu for Preservation Hold Library Review
            Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
            Write-Host "║         Preservation Hold Library Review Options           ║" -ForegroundColor Yellow
            Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
            Write-Host "║    1. Using Interactive Login (Delegated)                  ║" -ForegroundColor Cyan
            Write-Host "║    2. Using Certificate Auth (Application)                 ║" -ForegroundColor Cyan
            Write-Host "║    3. Authentication Information                           ║" -ForegroundColor Cyan
            Write-Host "║    4. Return to Main Menu                                  ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    
            $subChoice = Read-Host "`nEnter your choice (1-4)"
            
            switch ($subChoice) {
                "1" {
                    Get-PreservationHoldLibraryStatus
                }
                "2" {
                    Get-PreservationHoldLibraryStatus-App
                }
                "3" {
                    Write-Host "`nPnP PowerShell Authentication Documentation" -ForegroundColor Yellow
                    Write-Host "For detailed information about authentication methods, please visit:"
                    Write-Host "https://pnp.github.io/powershell/articles/authentication.html" -ForegroundColor Cyan
                    Write-Host "`nPress any key to return to main menu..."
                    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
                }
                "4" {
                    Write-Host "Returning to main menu..." -ForegroundColor Yellow
                }
                default {
                    Write-Host "Invalid selection. Returning to main menu..." -ForegroundColor Yellow
                }
            }
        }
        "8" {
            # Create submenu for Download Options
            Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Yellow
            Write-Host "║         Preservation Hold Library Download Options         ║" -ForegroundColor Yellow
            Write-Host "╠════════════════════════════════════════════════════════════╣" -ForegroundColor Yellow
            Write-Host "║    1. Using Interactive Login (Delegated)                  ║" -ForegroundColor Cyan
            Write-Host "║    2. Using Certificate Auth (Application)                 ║" -ForegroundColor Cyan
            Write-Host "║    3. Authentication Information                           ║" -ForegroundColor Cyan
            Write-Host "║    4. Return to Main Menu                                  ║" -ForegroundColor Cyan
            Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Yellow

            $subChoice = Read-Host "`nEnter your choice (1-4)"
        
            if ($subChoice -in "1", "2") {
                $selectedSites = Get-SiteSelection
                if ($selectedSites) {
                    foreach ($site in $selectedSites) {
                        # Create Downloads directory if it doesn't exist
                        $downloadsPath = ".\PreservationHoldFiles"
                        if (-not (Test-Path -Path $downloadsPath)) {
                            try {
                                New-Item -Path $downloadsPath -ItemType Directory -Force | Out-Null
                            }
                            catch {
                                Write-Host "Error creating Downloads directory: $($_.Exception.Message)" -ForegroundColor Red
                                throw
                            }
                        }
                    
                        # Create site-specific subfolder
                        $siteFolder = $site.Url -replace 'https://.+?.sharepoint.com', '' -replace '[^a-zA-Z0-9]', '_'
                        $downloadPath = Join-Path $downloadsPath $siteFolder
                    
                        switch ($subChoice) {
                            "1" {
                                DownloadPreservationHoldLibrary -SiteUrl $site.Url -DownloadPath $downloadPath
                            }
                            "2" {
                                DownloadPreservationHoldLibrary-App -SiteUrl $site.Url -DownloadPath $downloadPath
                            }
                        }
                    }
                }
            }
            elseif ($subChoice -eq "3") {
                Write-Host "`nPnP PowerShell Authentication Documentation" -ForegroundColor Yellow
                Write-Host "For detailed information about authentication methods, please visit:"
                Write-Host "https://pnp.github.io/powershell/articles/authentication.html" -ForegroundColor Cyan
                Write-Host "`nPress any key to return to main menu..."
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            }
            else {
                Write-Host "Returning to main menu..." -ForegroundColor Yellow
            }
        }
    }
}
catch {
    Write-Host "`nAn error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ((Get-PSSession).Name -like "SpoSession*") {
        Disconnect-SPOService
        Write-Host "`nDisconnected from SharePoint Online" -ForegroundColor Yellow
    }
    Stop-Transcript
    Write-Host "`nScript complete. Press any key to exit..." -ForegroundColor Green
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}