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
    
    $module = Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell"
    if (-not $module) {
        Write-Host "Installing SharePoint Online Management Shell..." -ForegroundColor Yellow
        Install-Module -Name "Microsoft.Online.SharePoint.PowerShell" -Force -AllowClobber
    }
    
    Import-Module "Microsoft.Online.SharePoint.PowerShell" -DisableNameChecking
    
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
            Write-Host "$index. $($_.Url)" -ForegroundColor White
            $sites += [PSCustomObject]@{
                Index = $index
                Url = $_.Url
                Title = $_.Title
            }
            $index++
        }

        do {
            Write-Host "`nSelect sites (comma-separated numbers, 'all' for all sites, or 'q' to quit):" -ForegroundColor Yellow
            $selection = Read-Host

            if ($selection -eq 'q') {
                Write-Host "Exiting script" -ForegroundColor Yellow
                return $null
            }

            $selectedSites = @()

            if ($selection -eq 'all') {
                $selectedSites = $sites
            } else {
                $selectedIndices = $selection -split ',' | ForEach-Object { $_.Trim() }
                $invalidIndices = @()

                foreach ($idx in $selectedIndices) {
                    if ([int]::TryParse($idx, [ref]$null)) {
                        $site = $sites | Where-Object { $_.Index -eq [int]$idx }
                        if ($site) {
                            $selectedSites += $site
                        } else {
                            $invalidIndices += $idx
                        }
                    } else {
                        $invalidIndices += $idx
                    }
                }

                if ($invalidIndices) {
                    Write-Host "`nWarning: Invalid site numbers detected: $($invalidIndices -join ', ')" -ForegroundColor Yellow
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
                Write-Host "- $($_.Url)" -ForegroundColor White
            }
            Write-Host "`nTotal sites selected: $($selectedSites.Count)" -ForegroundColor Cyan

            $confirm = Read-Host "`nProceed with these sites? (y/n)"
            if ($confirm -eq 'y') {
                return $selectedSites
            } else {
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
    
    # Show menu first
    Write-Host "`nSharePoint Version Management Options:" -ForegroundColor Yellow
    Write-Host "1. Enable Intelligent Versioning for tenant (new SPO sites)" -ForegroundColor Cyan
    Write-Host "2. Clean up existing versions by age" -ForegroundColor Cyan
    Write-Host "3. Clean up existing versions by count" -ForegroundColor Cyan
    Write-Host "4. Both enable Intelligent Versioning and clean up existing" -ForegroundColor Cyan
    Write-Host "5. Enable Intelligent Versioning for both tenant and existing SPO sites" -ForegroundColor Cyan
    Write-Host "6. Exit" -ForegroundColor Cyan
    
    $choice = Read-Host "`nEnter your choice (1-6)"
    
    # Exit early if user chooses to quit
    if ($choice -eq "6") {
        Write-Host "Exiting script" -ForegroundColor Yellow
        return
    }
    
    # Validate choice before connecting
    if ($choice -notin "1", "2", "3", "4", "5") {
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