#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows Update Script with GUI
.DESCRIPTION
    A PowerShell script with a simple UI to check and install:
    - Windows Updates (including optional driver updates)
    - Application updates via winget
.NOTES
    Run as Administrator for full functionality
#>

# Hide the PowerShell console window
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'
$consolePtr = [Console.Window]::GetConsoleWindow()
[Console.Window]::ShowWindow($consolePtr, 0) | Out-Null

# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Global variables
$script:WindowsUpdates = @()
$script:WingetUpdates = @()
$script:UpdateSession = $null
$script:UpdateSearcher = $null

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Windows System Updater"
$form.Size = New-Object System.Drawing.Size(800, 650)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows System Updater"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size = New-Object System.Drawing.Size(400, 40)
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$form.Controls.Add($titleLabel)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready. Click 'Check for Updates' to begin."
$statusLabel.Size = New-Object System.Drawing.Size(740, 25)
$statusLabel.Location = New-Object System.Drawing.Point(20, 55)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$form.Controls.Add($statusLabel)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(740, 23)
$progressBar.Location = New-Object System.Drawing.Point(20, 85)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Windows Updates Group Box
$winUpdatesGroup = New-Object System.Windows.Forms.GroupBox
$winUpdatesGroup.Text = "Windows Updates (including Optional Drivers)"
$winUpdatesGroup.Size = New-Object System.Drawing.Size(740, 180)
$winUpdatesGroup.Location = New-Object System.Drawing.Point(20, 120)
$form.Controls.Add($winUpdatesGroup)

# Windows Updates ListView
$winUpdatesList = New-Object System.Windows.Forms.ListView
$winUpdatesList.Size = New-Object System.Drawing.Size(720, 145)
$winUpdatesList.Location = New-Object System.Drawing.Point(10, 20)
$winUpdatesList.View = "Details"
$winUpdatesList.FullRowSelect = $true
$winUpdatesList.CheckBoxes = $true
$winUpdatesList.Columns.Add("Title", 450) | Out-Null
$winUpdatesList.Columns.Add("Size", 80) | Out-Null
$winUpdatesList.Columns.Add("Type", 100) | Out-Null
$winUpdatesList.Columns.Add("Status", 80) | Out-Null
$winUpdatesGroup.Controls.Add($winUpdatesList)

# Winget Updates Group Box
$wingetGroup = New-Object System.Windows.Forms.GroupBox
$wingetGroup.Text = "Application Updates (winget)"
$wingetGroup.Size = New-Object System.Drawing.Size(740, 180)
$wingetGroup.Location = New-Object System.Drawing.Point(20, 310)
$form.Controls.Add($wingetGroup)

# Winget Updates ListView
$wingetList = New-Object System.Windows.Forms.ListView
$wingetList.Size = New-Object System.Drawing.Size(720, 145)
$wingetList.Location = New-Object System.Drawing.Point(10, 20)
$wingetList.View = "Details"
$wingetList.FullRowSelect = $true
$wingetList.CheckBoxes = $true
$wingetList.Columns.Add("Application", 300) | Out-Null
$wingetList.Columns.Add("Current Version", 130) | Out-Null
$wingetList.Columns.Add("Available Version", 130) | Out-Null
$wingetList.Columns.Add("Status", 80) | Out-Null
$wingetGroup.Controls.Add($wingetList)

# Buttons Panel
$buttonPanel = New-Object System.Windows.Forms.Panel
$buttonPanel.Size = New-Object System.Drawing.Size(740, 50)
$buttonPanel.Location = New-Object System.Drawing.Point(20, 500)
$form.Controls.Add($buttonPanel)

# Check for Updates Button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Text = "Check for Updates"
$checkButton.Size = New-Object System.Drawing.Size(180, 40)
$checkButton.Location = New-Object System.Drawing.Point(0, 5)
$checkButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$checkButton.ForeColor = [System.Drawing.Color]::White
$checkButton.FlatStyle = "Flat"
$checkButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonPanel.Controls.Add($checkButton)

# Select All Button
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Size = New-Object System.Drawing.Size(120, 40)
$selectAllButton.Location = New-Object System.Drawing.Point(200, 5)
$selectAllButton.FlatStyle = "Flat"
$selectAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonPanel.Controls.Add($selectAllButton)

# Start Updates Button
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start Updates"
$startButton.Size = New-Object System.Drawing.Size(180, 40)
$startButton.Location = New-Object System.Drawing.Point(340, 5)
$startButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = "Flat"
$startButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$startButton.Enabled = $false
$buttonPanel.Controls.Add($startButton)

# Clean Temp Button
$cleanButton = New-Object System.Windows.Forms.Button
$cleanButton.Text = "Clean Temp Files"
$cleanButton.Size = New-Object System.Drawing.Size(140, 40)
$cleanButton.Location = New-Object System.Drawing.Point(540, 5)
$cleanButton.BackColor = [System.Drawing.Color]::FromArgb(180, 100, 20)
$cleanButton.ForeColor = [System.Drawing.Color]::White
$cleanButton.FlatStyle = "Flat"
$cleanButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$buttonPanel.Controls.Add($cleanButton)

# Reboot Button
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot Now"
$rebootButton.Size = New-Object System.Drawing.Size(120, 40)
$rebootButton.Location = New-Object System.Drawing.Point(540, 5)
$rebootButton.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 80)
$rebootButton.ForeColor = [System.Drawing.Color]::White
$rebootButton.FlatStyle = "Flat"
$rebootButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$rebootButton.Visible = $false
$buttonPanel.Controls.Add($rebootButton)

# Log TextBox
$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Log"
$logGroup.Size = New-Object System.Drawing.Size(740, 80)
$logGroup.Location = New-Object System.Drawing.Point(20, 555)
$form.Controls.Add($logGroup)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Size = New-Object System.Drawing.Size(720, 50)
$logBox.Location = New-Object System.Drawing.Point(10, 20)
$logBox.ReadOnly = $true
$logBox.BackColor = [System.Drawing.Color]::White
$logGroup.Controls.Add($logBox)

# Helper function to update log
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $logBox.AppendText("[$timestamp] $Message`r`n")
    $logBox.SelectionStart = $logBox.Text.Length
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# Helper function to update status
function Update-Status {
    param([string]$Message)
    $statusLabel.Text = $Message
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to format file size
function Format-FileSize {
    param([long]$Size)
    if ($Size -ge 1GB) { return "{0:N2} GB" -f ($Size / 1GB) }
    elseif ($Size -ge 1MB) { return "{0:N2} MB" -f ($Size / 1MB) }
    elseif ($Size -ge 1KB) { return "{0:N2} KB" -f ($Size / 1KB) }
    else { return "$Size B" }
}

# Function to check Windows Updates
function Get-WindowsUpdatesAvailable {
    Write-Log "Checking for Windows Updates..."
    Update-Status "Searching for Windows Updates (including optional drivers)..."
    
    try {
        $script:UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $script:UpdateSearcher = $script:UpdateSession.CreateUpdateSearcher()
        
        # Include driver updates and optional updates
        # Criteria: IsInstalled=0 means not installed, IsHidden=0 means not hidden
        $searchResult = $script:UpdateSearcher.Search("IsInstalled=0 and IsHidden=0")
        
        $script:WindowsUpdates = @()
        $winUpdatesList.Items.Clear()
        
        foreach ($update in $searchResult.Updates) {
            $script:WindowsUpdates += $update
            
            $item = New-Object System.Windows.Forms.ListViewItem($update.Title)
            $item.Checked = $true
            
            # Get size
            $size = if ($update.MaxDownloadSize -gt 0) { Format-FileSize $update.MaxDownloadSize } else { "Unknown" }
            $item.SubItems.Add($size) | Out-Null
            
            # Determine type
            $type = "Update"
            if ($update.DriverClass) { $type = "Driver" }
            elseif ($update.Title -match "Cumulative|Security") { $type = "Security" }
            elseif ($update.Title -match "Feature") { $type = "Feature" }
            $item.SubItems.Add($type) | Out-Null
            
            $item.SubItems.Add("Pending") | Out-Null
            $item.Tag = $update
            
            $winUpdatesList.Items.Add($item) | Out-Null
        }
        
        $count = $script:WindowsUpdates.Count
        Write-Log "Found $count Windows Update(s)"
        return $count
    }
    catch {
        Write-Log "Error checking Windows Updates: $_"
        return 0
    }
}

# Function to check Winget Updates
function Get-WingetUpdatesAvailable {
    Write-Log "Checking for application updates via winget..."
    Update-Status "Searching for application updates..."
    
    try {
        # Check if winget is available
        $wingetPath = Get-Command winget -ErrorAction SilentlyContinue
        if (-not $wingetPath) {
            Write-Log "Winget is not installed or not in PATH"
            return 0
        }
        
        # Get upgradable packages
        $wingetOutput = winget upgrade --include-unknown 2>&1 | Out-String
        
        $script:WingetUpdates = @()
        $wingetList.Items.Clear()
        
        # Parse winget output
        $lines = $wingetOutput -split "`n"
        $headerFound = $false
        $separatorFound = $false
        
        foreach ($line in $lines) {
            if ($line -match "^Name\s+Id\s+Version\s+Available") {
                $headerFound = $true
                continue
            }
            if ($headerFound -and $line -match "^-+") {
                $separatorFound = $true
                continue
            }
            if ($separatorFound -and $line.Trim() -ne "" -and $line -notmatch "upgrades available" -and $line -notmatch "^The following") {
                # Parse the line - winget uses fixed-width columns
                if ($line.Length -gt 20) {
                    # Try to parse the update info
                    $parts = $line -split '\s{2,}'
                    if ($parts.Count -ge 4) {
                        $appName = $parts[0].Trim()
                        $appId = $parts[1].Trim()
                        $currentVer = $parts[2].Trim()
                        $availableVer = $parts[3].Trim()
                        
                        if ($appName -and $availableVer -and $appName -notmatch "^-+$") {
                            $updateInfo = @{
                                Name = $appName
                                Id = $appId
                                CurrentVersion = $currentVer
                                AvailableVersion = $availableVer
                            }
                            $script:WingetUpdates += $updateInfo
                            
                            $item = New-Object System.Windows.Forms.ListViewItem($appName)
                            $item.Checked = $true
                            $item.SubItems.Add($currentVer) | Out-Null
                            $item.SubItems.Add($availableVer) | Out-Null
                            $item.SubItems.Add("Pending") | Out-Null
                            $item.Tag = $updateInfo
                            
                            $wingetList.Items.Add($item) | Out-Null
                        }
                    }
                }
            }
        }
        
        $count = $script:WingetUpdates.Count
        Write-Log "Found $count application update(s)"
        return $count
    }
    catch {
        Write-Log "Error checking winget updates: $_"
        return 0
    }
}

# Function to install Windows Updates
function Install-WindowsUpdatesSelected {
    $selectedUpdates = @()
    foreach ($item in $winUpdatesList.Items) {
        if ($item.Checked) {
            $selectedUpdates += $item.Tag
        }
    }
    
    if ($selectedUpdates.Count -eq 0) {
        Write-Log "No Windows Updates selected"
        return $false
    }
    
    Write-Log "Installing $($selectedUpdates.Count) Windows Update(s)..."
    
    try {
        $updatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in $selectedUpdates) {
            $updatesToDownload.Add($update) | Out-Null
        }
        
        # Download updates
        Update-Status "Downloading Windows Updates..."
        $downloader = $script:UpdateSession.CreateUpdateDownloader()
        $downloader.Updates = $updatesToDownload
        
        $progressBar.Value = 0
        $progressBar.Maximum = $selectedUpdates.Count * 2
        
        $downloadResult = $downloader.Download()
        
        if ($downloadResult.ResultCode -eq 2) {
            Write-Log "Downloads completed successfully"
        }
        
        # Update list items
        $itemIndex = 0
        foreach ($item in $winUpdatesList.Items) {
            if ($item.Checked) {
                $item.SubItems[3].Text = "Downloaded"
                $progressBar.Value = $itemIndex + 1
                $itemIndex++
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Install updates
        Update-Status "Installing Windows Updates..."
        $installer = $script:UpdateSession.CreateUpdateInstaller()
        $installer.Updates = $updatesToDownload
        
        $installResult = $installer.Install()
        
        $itemIndex = 0
        foreach ($item in $winUpdatesList.Items) {
            if ($item.Checked) {
                $resultCode = $installResult.GetUpdateResult($itemIndex).ResultCode
                $status = switch ($resultCode) {
                    2 { "Installed" }
                    3 { "Installed (Reboot)" }
                    4 { "Failed" }
                    5 { "Aborted" }
                    default { "Unknown" }
                }
                $item.SubItems[3].Text = $status
                $progressBar.Value = $selectedUpdates.Count + $itemIndex + 1
                $itemIndex++
            }
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        Write-Log "Windows Updates installation completed"
        return ($installResult.RebootRequired)
    }
    catch {
        Write-Log "Error installing Windows Updates: $_"
        return $false
    }
}

# Function to install Winget Updates
function Install-WingetUpdatesSelected {
    $selectedApps = @()
    foreach ($item in $wingetList.Items) {
        if ($item.Checked) {
            $selectedApps += @{
                Item = $item
                Info = $item.Tag
            }
        }
    }
    
    if ($selectedApps.Count -eq 0) {
        Write-Log "No application updates selected"
        return
    }
    
    Write-Log "Installing $($selectedApps.Count) application update(s)..."
    
    $current = 0
    foreach ($app in $selectedApps) {
        $current++
        $appInfo = $app.Info
        $item = $app.Item
        
        Update-Status "Updating: $($appInfo.Name) ($current/$($selectedApps.Count))..."
        $item.SubItems[3].Text = "Installing..."
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $result = winget upgrade --id $appInfo.Id --silent --accept-package-agreements --accept-source-agreements 2>&1
            
            if ($LASTEXITCODE -eq 0) {
                $item.SubItems[3].Text = "Updated"
                Write-Log "Updated: $($appInfo.Name)"
            }
            else {
                $item.SubItems[3].Text = "Failed"
                Write-Log "Failed to update: $($appInfo.Name)"
            }
        }
        catch {
            $item.SubItems[3].Text = "Error"
            Write-Log "Error updating $($appInfo.Name): $_"
        }
        
        $progressBar.Value = [math]::Min(100, [int](($current / $selectedApps.Count) * 100))
        [System.Windows.Forms.Application]::DoEvents()
    }
    
    Write-Log "Application updates completed"
}

# Function to clean temporary files
function Clear-TemporaryFiles {
    Write-Log "Starting temporary files cleanup..."
    Update-Status "Cleaning temporary files..."
    
    $totalFreed = 0
    $foldersToClean = @(
        @{ Path = $env:TEMP; Name = "User Temp" },
        @{ Path = "C:\Windows\Temp"; Name = "Windows Temp" },
        @{ Path = "C:\Windows\Prefetch"; Name = "Prefetch" },
        @{ Path = "C:\Windows\SoftwareDistribution\Download"; Name = "Windows Update Cache" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\INetCache"; Name = "IE Cache" },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer"; Name = "Explorer Thumbnails" },
        @{ Path = "$env:LOCALAPPDATA\Temp"; Name = "Local Temp" },
        @{ Path = "C:\Windows\Logs\CBS"; Name = "CBS Logs" },
        @{ Path = "$env:LOCALAPPDATA\CrashDumps"; Name = "Crash Dumps" },
        @{ Path = "C:\ProgramData\Microsoft\Windows\WER"; Name = "Error Reports" }
    )
    
    $progressBar.Value = 0
    $progressBar.Maximum = $foldersToClean.Count + 2
    $currentStep = 0
    
    foreach ($folder in $foldersToClean) {
        $currentStep++
        $progressBar.Value = $currentStep
        [System.Windows.Forms.Application]::DoEvents()
        
        if (Test-Path $folder.Path) {
            try {
                $sizeBefore = (Get-ChildItem -Path $folder.Path -Recurse -Force -ErrorAction SilentlyContinue | 
                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if (-not $sizeBefore) { $sizeBefore = 0 }
                
                # For thumbnails, only delete specific cache files
                if ($folder.Name -eq "Explorer Thumbnails") {
                    Get-ChildItem -Path $folder.Path -Filter "thumbcache_*.db" -Force -ErrorAction SilentlyContinue | 
                        Remove-Item -Force -ErrorAction SilentlyContinue
                }
                else {
                    Get-ChildItem -Path $folder.Path -Recurse -Force -ErrorAction SilentlyContinue | 
                        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
                }
                
                $sizeAfter = (Get-ChildItem -Path $folder.Path -Recurse -Force -ErrorAction SilentlyContinue | 
                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                if (-not $sizeAfter) { $sizeAfter = 0 }
                
                $freed = $sizeBefore - $sizeAfter
                if ($freed -gt 0) {
                    $totalFreed += $freed
                    Write-Log "$($folder.Name): Freed $(Format-FileSize $freed)"
                }
            }
            catch {
                Write-Log "Could not fully clean $($folder.Name)"
            }
        }
    }
    
    # Run Windows Disk Cleanup for WinSxS (component cleanup)
    $currentStep++
    $progressBar.Value = $currentStep
    Update-Status "Running Component Cleanup (WinSxS)..."
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        Write-Log "Running DISM Component Cleanup..."
        $dismResult = Start-Process -FilePath "dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup /ResetBase" -Wait -PassThru -NoNewWindow -ErrorAction SilentlyContinue
        if ($dismResult.ExitCode -eq 0) {
            Write-Log "Component Cleanup completed successfully"
        }
    }
    catch {
        Write-Log "Component Cleanup skipped or failed"
    }
    
    # Clear Windows Update cache service (stop, clean, start)
    $currentStep++
    $progressBar.Value = $currentStep
    Update-Status "Cleaning Windows Update service cache..."
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue
        Stop-Service -Name bits -Force -ErrorAction SilentlyContinue
        
        if (Test-Path "C:\Windows\SoftwareDistribution\DataStore") {
            Remove-Item -Path "C:\Windows\SoftwareDistribution\DataStore\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path "C:\Windows\SoftwareDistribution\Download") {
            Remove-Item -Path "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
        Write-Log "Windows Update cache cleaned"
    }
    catch {
        Write-Log "Could not fully clean Windows Update cache"
        Start-Service -Name bits -ErrorAction SilentlyContinue
        Start-Service -Name wuauserv -ErrorAction SilentlyContinue
    }
    
    $progressBar.Value = $progressBar.Maximum
    
    Write-Log "Cleanup completed! Total freed: $(Format-FileSize $totalFreed)"
    Update-Status "Cleanup completed! Freed approximately $(Format-FileSize $totalFreed)"
    
    [System.Windows.Forms.MessageBox]::Show(
        "Temporary files cleanup completed!`n`nFreed approximately $(Format-FileSize $totalFreed) of disk space.`n`nNote: Some files may have been in use and could not be deleted.",
        "Cleanup Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

# Check Button Click Event
$checkButton.Add_Click({
    $checkButton.Enabled = $false
    $startButton.Enabled = $false
    $progressBar.Style = "Marquee"
    
    $winCount = Get-WindowsUpdatesAvailable
    $appCount = Get-WingetUpdatesAvailable
    
    $progressBar.Style = "Continuous"
    $progressBar.Value = 0
    
    $totalUpdates = $winCount + $appCount
    if ($totalUpdates -gt 0) {
        Update-Status "Found $winCount Windows Update(s) and $appCount Application Update(s)"
        $startButton.Enabled = $true
    }
    else {
        Update-Status "No updates available. Your system is up to date!"
    }
    
    $checkButton.Enabled = $true
    Write-Log "Update check completed"
})

# Select All Button Click Event
$selectAllButton.Add_Click({
    $allChecked = $true
    foreach ($item in $winUpdatesList.Items) {
        if (-not $item.Checked) { $allChecked = $false; break }
    }
    foreach ($item in $wingetList.Items) {
        if (-not $item.Checked) { $allChecked = $false; break }
    }
    
    $newState = -not $allChecked
    foreach ($item in $winUpdatesList.Items) { $item.Checked = $newState }
    foreach ($item in $wingetList.Items) { $item.Checked = $newState }
    
    $selectAllButton.Text = if ($newState) { "Deselect All" } else { "Select All" }
})

# Start Updates Button Click Event
$startButton.Add_Click({
    $checkButton.Enabled = $false
    $startButton.Enabled = $false
    $selectAllButton.Enabled = $false
    
    $progressBar.Value = 0
    $progressBar.Maximum = 100
    
    Write-Log "Starting update installation..."
    
    # Install Windows Updates first
    $rebootRequired = Install-WindowsUpdatesSelected
    
    # Then install Winget updates
    Install-WingetUpdatesSelected
    
    Update-Status "All updates completed!"
    Write-Log "All update operations finished"
    
    $progressBar.Value = 100
    
    # Show reboot prompt if needed
    if ($rebootRequired) {
        $rebootButton.Visible = $true
        Update-Status "Updates completed. A reboot is required to finish installation."
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Updates have been installed successfully.`n`nA system restart is required to complete the installation.`n`nWould you like to restart now?",
            "Restart Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "User initiated system restart"
            Restart-Computer -Force
        }
    }
    else {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "All updates have been installed successfully!`n`nIt's recommended to restart your computer.`n`nWould you like to restart now?",
            "Updates Complete",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "User initiated system restart"
            Restart-Computer -Force
        }
        else {
            $rebootButton.Visible = $true
        }
    }
    
    $checkButton.Enabled = $true
    $selectAllButton.Enabled = $true
})

# Reboot Button Click Event
$rebootButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to restart your computer now?`n`nPlease save all your work before continuing.",
        "Confirm Restart",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Write-Log "Initiating system restart..."
        Restart-Computer -Force
    }
})

# Clean Temp Button Click Event
$cleanButton.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will clean temporary files including:`n`n- User and System Temp folders`n- Windows Update cache`n- Prefetch files`n- Thumbnail cache`n- Error reports and crash dumps`n- WinSxS component cleanup`n`nThis may take several minutes. Continue?",
        "Clean Temporary Files",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $cleanButton.Enabled = $false
        $checkButton.Enabled = $false
        $startButton.Enabled = $false
        
        Clear-TemporaryFiles
        
        $cleanButton.Enabled = $true
        $checkButton.Enabled = $true
    }
})

# Form closing event
$form.Add_FormClosing({
    $script:UpdateSession = $null
    $script:UpdateSearcher = $null
})

# Check for admin rights on startup
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    [System.Windows.Forms.MessageBox]::Show(
        "This application requires administrator privileges.`n`nPlease right-click the script and select 'Run as Administrator'.",
        "Administrator Required",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    $checkButton.Enabled = $false
    $statusLabel.Text = "[!] Please restart as Administrator"
    $statusLabel.ForeColor = [System.Drawing.Color]::Red
}

Write-Log "Windows System Updater started"

# Show the form
[void]$form.ShowDialog()
