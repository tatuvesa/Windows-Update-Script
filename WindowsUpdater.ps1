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

# Script path for auto-restart feature
$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:AutoRestartTaskName = "WindowsUpdaterAutoRestart"
$script:DarkMode = $false

# Check if this is an auto-restart run
$script:IsAutoRestart = $false
if (Test-Path "HKCU:\Software\WindowsUpdater") {
    $autoRestartFlag = Get-ItemProperty -Path "HKCU:\Software\WindowsUpdater" -Name "AutoRestart" -ErrorAction SilentlyContinue
    if ($autoRestartFlag.AutoRestart -eq 1) {
        $script:IsAutoRestart = $true
        # Clean up the flag
        Remove-ItemProperty -Path "HKCU:\Software\WindowsUpdater" -Name "AutoRestart" -ErrorAction SilentlyContinue
        # Remove the scheduled task
        Unregister-ScheduledTask -TaskName $script:AutoRestartTaskName -Confirm:$false -ErrorAction SilentlyContinue
    }
}

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
$form.Size = New-Object System.Drawing.Size(850, 720)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedSingle"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

# Color schemes
$script:LightColors = @{
    FormBack  = [System.Drawing.Color]::FromArgb(240, 240, 240)
    GroupBack = [System.Drawing.Color]::FromArgb(240, 240, 240)
    TextColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
    ListBack  = [System.Drawing.Color]::White
    LogBack   = [System.Drawing.Color]::White
}
$script:DarkColors = @{
    FormBack  = [System.Drawing.Color]::FromArgb(32, 32, 32)
    GroupBack = [System.Drawing.Color]::FromArgb(45, 45, 45)
    TextColor = [System.Drawing.Color]::FromArgb(220, 220, 220)
    ListBack  = [System.Drawing.Color]::FromArgb(45, 45, 45)
    LogBack   = [System.Drawing.Color]::FromArgb(30, 30, 30)
}

# Title Label
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Windows System Updater"
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$titleLabel.Size = New-Object System.Drawing.Size(350, 40)
$titleLabel.Location = New-Object System.Drawing.Point(20, 15)
$form.Controls.Add($titleLabel)

# Disk Space Label
$diskSpaceLabel = New-Object System.Windows.Forms.Label
$diskSpaceLabel.Text = "Disk: --"
$diskSpaceLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$diskSpaceLabel.Size = New-Object System.Drawing.Size(200, 20)
$diskSpaceLabel.Location = New-Object System.Drawing.Point(410, 15)
$diskSpaceLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($diskSpaceLabel)

# Network Status Label
$networkLabel = New-Object System.Windows.Forms.Label
$networkLabel.Text = "Network: Checking..."
$networkLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$networkLabel.Size = New-Object System.Drawing.Size(200, 20)
$networkLabel.Location = New-Object System.Drawing.Point(410, 35)
$networkLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$form.Controls.Add($networkLabel)

# Dark Mode Toggle Button
$darkModeButton = New-Object System.Windows.Forms.Button
$darkModeButton.Text = "Dark Mode"
$darkModeButton.Size = New-Object System.Drawing.Size(100, 30)
$darkModeButton.Location = New-Object System.Drawing.Point(620, 20)
$darkModeButton.FlatStyle = "Flat"
$darkModeButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($darkModeButton)

# History Button
$historyButton = New-Object System.Windows.Forms.Button
$historyButton.Text = "Update History"
$historyButton.Size = New-Object System.Drawing.Size(100, 30)
$historyButton.Location = New-Object System.Drawing.Point(725, 20)
$historyButton.FlatStyle = "Flat"
$historyButton.Font = New-Object System.Drawing.Font("Segoe UI", 8)
$form.Controls.Add($historyButton)

# Status Label
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Ready. Click 'Check for Updates' to begin."
$statusLabel.Size = New-Object System.Drawing.Size(790, 25)
$statusLabel.Location = New-Object System.Drawing.Point(20, 55)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(60, 60, 60)
$form.Controls.Add($statusLabel)

# Progress Bar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Size = New-Object System.Drawing.Size(790, 23)
$progressBar.Location = New-Object System.Drawing.Point(20, 85)
$progressBar.Style = "Continuous"
$form.Controls.Add($progressBar)

# Windows Updates Group Box
$winUpdatesGroup = New-Object System.Windows.Forms.GroupBox
$winUpdatesGroup.Text = "Windows Updates (including Optional Drivers)"
$winUpdatesGroup.Size = New-Object System.Drawing.Size(790, 180)
$winUpdatesGroup.Location = New-Object System.Drawing.Point(20, 120)
$form.Controls.Add($winUpdatesGroup)

# Windows Updates ListView
$winUpdatesList = New-Object System.Windows.Forms.ListView
$winUpdatesList.Size = New-Object System.Drawing.Size(770, 145)
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
$wingetGroup.Size = New-Object System.Drawing.Size(790, 180)
$wingetGroup.Location = New-Object System.Drawing.Point(20, 310)
$form.Controls.Add($wingetGroup)

# Winget Updates ListView
$wingetList = New-Object System.Windows.Forms.ListView
$wingetList.Size = New-Object System.Drawing.Size(770, 145)
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
$buttonPanel.Size = New-Object System.Drawing.Size(790, 90)
$buttonPanel.Location = New-Object System.Drawing.Point(20, 500)
$form.Controls.Add($buttonPanel)

# Check for Updates Button
$checkButton = New-Object System.Windows.Forms.Button
$checkButton.Text = "Check for Updates"
$checkButton.Size = New-Object System.Drawing.Size(170, 40)
$checkButton.Location = New-Object System.Drawing.Point(0, 5)
$checkButton.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$checkButton.ForeColor = [System.Drawing.Color]::White
$checkButton.FlatStyle = "Flat"
$checkButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$buttonPanel.Controls.Add($checkButton)

# Select All Button
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Size = New-Object System.Drawing.Size(100, 40)
$selectAllButton.Location = New-Object System.Drawing.Point(180, 5)
$selectAllButton.FlatStyle = "Flat"
$selectAllButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$buttonPanel.Controls.Add($selectAllButton)

# Start Updates Button
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "Start Updates"
$startButton.Size = New-Object System.Drawing.Size(150, 40)
$startButton.Location = New-Object System.Drawing.Point(290, 5)
$startButton.BackColor = [System.Drawing.Color]::FromArgb(16, 124, 16)
$startButton.ForeColor = [System.Drawing.Color]::White
$startButton.FlatStyle = "Flat"
$startButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$startButton.Enabled = $false
$buttonPanel.Controls.Add($startButton)

# Clean Temp Button
$cleanButton = New-Object System.Windows.Forms.Button
$cleanButton.Text = "Clean Temp"
$cleanButton.Size = New-Object System.Drawing.Size(110, 40)
$cleanButton.Location = New-Object System.Drawing.Point(450, 5)
$cleanButton.BackColor = [System.Drawing.Color]::FromArgb(180, 100, 20)
$cleanButton.ForeColor = [System.Drawing.Color]::White
$cleanButton.FlatStyle = "Flat"
$cleanButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$buttonPanel.Controls.Add($cleanButton)

# Reboot Button
$rebootButton = New-Object System.Windows.Forms.Button
$rebootButton.Text = "Reboot Now"
$rebootButton.Size = New-Object System.Drawing.Size(110, 40)
$rebootButton.Location = New-Object System.Drawing.Point(570, 5)
$rebootButton.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 80)
$rebootButton.ForeColor = [System.Drawing.Color]::White
$rebootButton.FlatStyle = "Flat"
$rebootButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$rebootButton.Visible = $false
$buttonPanel.Controls.Add($rebootButton)

# Auto-Restart Checkbox
$autoRestartCheck = New-Object System.Windows.Forms.CheckBox
$autoRestartCheck.Text = "Relaunch after restart"
$autoRestartCheck.Size = New-Object System.Drawing.Size(180, 25)
$autoRestartCheck.Location = New-Object System.Drawing.Point(0, 50)
$autoRestartCheck.Checked = $true
$buttonPanel.Controls.Add($autoRestartCheck)

# Scan Leftovers Button
$leftoversButton = New-Object System.Windows.Forms.Button
$leftoversButton.Text = "Scan Leftovers"
$leftoversButton.Size = New-Object System.Drawing.Size(110, 25)
$leftoversButton.Location = New-Object System.Drawing.Point(200, 50)
$leftoversButton.BackColor = [System.Drawing.Color]::FromArgb(100, 100, 160)
$leftoversButton.ForeColor = [System.Drawing.Color]::White
$leftoversButton.FlatStyle = "Flat"
$leftoversButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$buttonPanel.Controls.Add($leftoversButton)

# Reboot with Auto-Start Button (hidden initially)
$rebootAutoButton = New-Object System.Windows.Forms.Button
$rebootAutoButton.Text = "Reboot + Continue"
$rebootAutoButton.Size = New-Object System.Drawing.Size(130, 40)
$rebootAutoButton.Location = New-Object System.Drawing.Point(690, 5)
$rebootAutoButton.BackColor = [System.Drawing.Color]::FromArgb(130, 60, 60)
$rebootAutoButton.ForeColor = [System.Drawing.Color]::White
$rebootAutoButton.FlatStyle = "Flat"
$rebootAutoButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$rebootAutoButton.Visible = $false
$buttonPanel.Controls.Add($rebootAutoButton)

# Log TextBox
$logGroup = New-Object System.Windows.Forms.GroupBox
$logGroup.Text = "Log"
$logGroup.Size = New-Object System.Drawing.Size(790, 100)
$logGroup.Location = New-Object System.Drawing.Point(20, 600)
$form.Controls.Add($logGroup)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.Size = New-Object System.Drawing.Size(770, 70)
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

# Function to update disk space display
function Update-DiskSpace {
    try {
        $drive = Get-PSDrive -Name C -ErrorAction SilentlyContinue
        if ($drive) {
            $free = Format-FileSize ($drive.Free)
            $used = Format-FileSize ($drive.Used)
            $total = Format-FileSize ($drive.Free + $drive.Used)
            $diskSpaceLabel.Text = "Disk C: $free free / $total"
        }
    }
    catch {
        $diskSpaceLabel.Text = "Disk: Unable to read"
    }
}

# Function to check network connectivity
function Test-NetworkConnection {
    $networkLabel.Text = "Network: Checking..."
    [System.Windows.Forms.Application]::DoEvents()
    
    try {
        $result = Test-Connection -ComputerName "www.microsoft.com" -Count 1 -Quiet -ErrorAction SilentlyContinue
        if ($result) {
            $networkLabel.Text = "Network: Connected"
            $networkLabel.ForeColor = if ($script:DarkMode) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Green }
            return $true
        }
        else {
            $networkLabel.Text = "Network: No Connection"
            $networkLabel.ForeColor = [System.Drawing.Color]::Red
            return $false
        }
    }
    catch {
        $networkLabel.Text = "Network: Error"
        $networkLabel.ForeColor = [System.Drawing.Color]::Red
        return $false
    }
}

# Function to toggle dark mode
function Set-DarkMode {
    param([bool]$Enable)
    
    $script:DarkMode = $Enable
    $colors = if ($Enable) { $script:DarkColors } else { $script:LightColors }
    
    # Form
    $form.BackColor = $colors.FormBack
    
    # Labels
    $titleLabel.ForeColor = $colors.TextColor
    $statusLabel.ForeColor = $colors.TextColor
    $diskSpaceLabel.ForeColor = $colors.TextColor
    $autoRestartCheck.ForeColor = $colors.TextColor
    $autoRestartCheck.BackColor = $colors.FormBack
    
    # Group boxes
    $winUpdatesGroup.ForeColor = $colors.TextColor
    $winUpdatesGroup.BackColor = $colors.FormBack
    $wingetGroup.ForeColor = $colors.TextColor
    $wingetGroup.BackColor = $colors.FormBack
    $logGroup.ForeColor = $colors.TextColor
    $logGroup.BackColor = $colors.FormBack
    
    # ListViews
    $winUpdatesList.BackColor = $colors.ListBack
    $winUpdatesList.ForeColor = $colors.TextColor
    $wingetList.BackColor = $colors.ListBack
    $wingetList.ForeColor = $colors.TextColor
    
    # Log
    $logBox.BackColor = $colors.LogBack
    $logBox.ForeColor = $colors.TextColor
    
    # Button panel
    $buttonPanel.BackColor = $colors.FormBack
    
    # Update button text
    $darkModeButton.Text = if ($Enable) { "Light Mode" } else { "Dark Mode" }
    
    # Re-check network color
    if ($networkLabel.Text -match "Connected") {
        $networkLabel.ForeColor = if ($Enable) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Green }
    }
}

# Function to register auto-restart
function Register-AutoRestart {
    try {
        # Create registry key to flag auto-restart
        if (-not (Test-Path "HKCU:\Software\WindowsUpdater")) {
            New-Item -Path "HKCU:\Software\WindowsUpdater" -Force | Out-Null
        }
        Set-ItemProperty -Path "HKCU:\Software\WindowsUpdater" -Name "AutoRestart" -Value 1 -Type DWord
        
        # Get the batch file path (same directory as script)
        $scriptDir = Split-Path -Parent $script:ScriptPath
        $batchPath = Join-Path $scriptDir "RunUpdater.bat"
        
        # Create scheduled task to run at logon
        $action = New-ScheduledTaskAction -Execute $batchPath -WorkingDirectory $scriptDir
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -RunLevel Highest
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        
        Register-ScheduledTask -TaskName $script:AutoRestartTaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
        
        Write-Log "Auto-restart registered successfully"
        return $true
    }
    catch {
        Write-Log "Failed to register auto-restart: $_"
        return $false
    }
}

# Function to show update history
function Show-UpdateHistory {
    $historyForm = New-Object System.Windows.Forms.Form
    $historyForm.Text = "Windows Update History"
    $historyForm.Size = New-Object System.Drawing.Size(700, 500)
    $historyForm.StartPosition = "CenterParent"
    $historyForm.FormBorderStyle = "FixedDialog"
    $historyForm.MaximizeBox = $false
    $historyForm.MinimizeBox = $false
    
    if ($script:DarkMode) {
        $historyForm.BackColor = $script:DarkColors.FormBack
    }
    
    $historyList = New-Object System.Windows.Forms.ListView
    $historyList.Size = New-Object System.Drawing.Size(660, 420)
    $historyList.Location = New-Object System.Drawing.Point(10, 10)
    $historyList.View = "Details"
    $historyList.FullRowSelect = $true
    $historyList.Columns.Add("Date", 120) | Out-Null
    $historyList.Columns.Add("Title", 350) | Out-Null
    $historyList.Columns.Add("Status", 80) | Out-Null
    $historyList.Columns.Add("Type", 80) | Out-Null
    
    if ($script:DarkMode) {
        $historyList.BackColor = $script:DarkColors.ListBack
        $historyList.ForeColor = $script:DarkColors.TextColor
    }
    
    $historyForm.Controls.Add($historyList)
    
    # Get update history
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $historyCount = $searcher.GetTotalHistoryCount()
        $history = $searcher.QueryHistory(0, [Math]::Min($historyCount, 50))
        
        foreach ($entry in $history) {
            if ($entry.Title) {
                $item = New-Object System.Windows.Forms.ListViewItem($entry.Date.ToString("yyyy-MM-dd HH:mm"))
                $item.SubItems.Add($entry.Title) | Out-Null
                
                $status = switch ($entry.ResultCode) {
                    1 { "In Progress" }
                    2 { "Succeeded" }
                    3 { "Succeeded (Errors)" }
                    4 { "Failed" }
                    5 { "Aborted" }
                    default { "Unknown" }
                }
                $item.SubItems.Add($status) | Out-Null
                
                $type = switch ($entry.Operation) {
                    1 { "Install" }
                    2 { "Uninstall" }
                    default { "Other" }
                }
                $item.SubItems.Add($type) | Out-Null
                
                # Color code by status
                if ($entry.ResultCode -eq 4) {
                    $item.ForeColor = [System.Drawing.Color]::Red
                }
                elseif ($entry.ResultCode -eq 2) {
                    $item.ForeColor = if ($script:DarkMode) { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Green }
                }
                
                $historyList.Items.Add($item) | Out-Null
            }
        }
    }
    catch {
        $item = New-Object System.Windows.Forms.ListViewItem("Error loading history")
        $historyList.Items.Add($item) | Out-Null
    }
    
    [void]$historyForm.ShowDialog()
}

# Function to get installed programs
function Get-InstalledPrograms {
    $programs = @()
    
    # Registry paths for installed programs
    $regPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($path in $regPaths) {
        try {
            $items = Get-ItemProperty $path -ErrorAction SilentlyContinue
            foreach ($item in $items) {
                if ($item.DisplayName) {
                    $programs += @{
                        Name = $item.DisplayName
                        Publisher = $item.Publisher
                        InstallLocation = $item.InstallLocation
                        UninstallString = $item.UninstallString
                    }
                }
            }
        }
        catch { }
    }
    
    return $programs
}

# Function to scan for leftover files
function Show-LeftoverScanner {
    $scanForm = New-Object System.Windows.Forms.Form
    $scanForm.Text = "Leftover Files Scanner"
    $scanForm.Size = New-Object System.Drawing.Size(800, 600)
    $scanForm.StartPosition = "CenterParent"
    $scanForm.FormBorderStyle = "FixedDialog"
    $scanForm.MaximizeBox = $false
    $scanForm.MinimizeBox = $false
    
    if ($script:DarkMode) {
        $scanForm.BackColor = $script:DarkColors.FormBack
    }
    
    # Status label
    $scanStatus = New-Object System.Windows.Forms.Label
    $scanStatus.Text = "Click 'Scan' to search for leftover files from uninstalled programs..."
    $scanStatus.Size = New-Object System.Drawing.Size(760, 25)
    $scanStatus.Location = New-Object System.Drawing.Point(10, 10)
    if ($script:DarkMode) { $scanStatus.ForeColor = $script:DarkColors.TextColor }
    $scanForm.Controls.Add($scanStatus)
    
    # Progress bar
    $scanProgress = New-Object System.Windows.Forms.ProgressBar
    $scanProgress.Size = New-Object System.Drawing.Size(760, 20)
    $scanProgress.Location = New-Object System.Drawing.Point(10, 40)
    $scanForm.Controls.Add($scanProgress)
    
    # Results ListView
    $leftoversList = New-Object System.Windows.Forms.ListView
    $leftoversList.Size = New-Object System.Drawing.Size(760, 400)
    $leftoversList.Location = New-Object System.Drawing.Point(10, 70)
    $leftoversList.View = "Details"
    $leftoversList.FullRowSelect = $true
    $leftoversList.CheckBoxes = $true
    $leftoversList.Columns.Add("Folder/File", 350) | Out-Null
    $leftoversList.Columns.Add("Location", 200) | Out-Null
    $leftoversList.Columns.Add("Size", 80) | Out-Null
    $leftoversList.Columns.Add("Type", 100) | Out-Null
    
    if ($script:DarkMode) {
        $leftoversList.BackColor = $script:DarkColors.ListBack
        $leftoversList.ForeColor = $script:DarkColors.TextColor
    }
    $scanForm.Controls.Add($leftoversList)
    
    # Buttons panel
    $scanButtonPanel = New-Object System.Windows.Forms.Panel
    $scanButtonPanel.Size = New-Object System.Drawing.Size(760, 40)
    $scanButtonPanel.Location = New-Object System.Drawing.Point(10, 480)
    $scanForm.Controls.Add($scanButtonPanel)
    
    # Scan button
    $scanBtn = New-Object System.Windows.Forms.Button
    $scanBtn.Text = "Scan for Leftovers"
    $scanBtn.Size = New-Object System.Drawing.Size(150, 35)
    $scanBtn.Location = New-Object System.Drawing.Point(0, 0)
    $scanBtn.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
    $scanBtn.ForeColor = [System.Drawing.Color]::White
    $scanBtn.FlatStyle = "Flat"
    $scanButtonPanel.Controls.Add($scanBtn)
    
    # Select All button
    $selectAllBtn = New-Object System.Windows.Forms.Button
    $selectAllBtn.Text = "Select All"
    $selectAllBtn.Size = New-Object System.Drawing.Size(100, 35)
    $selectAllBtn.Location = New-Object System.Drawing.Point(160, 0)
    $selectAllBtn.FlatStyle = "Flat"
    $scanButtonPanel.Controls.Add($selectAllBtn)
    
    # Delete Selected button
    $deleteBtn = New-Object System.Windows.Forms.Button
    $deleteBtn.Text = "Delete Selected"
    $deleteBtn.Size = New-Object System.Drawing.Size(130, 35)
    $deleteBtn.Location = New-Object System.Drawing.Point(270, 0)
    $deleteBtn.BackColor = [System.Drawing.Color]::FromArgb(200, 80, 80)
    $deleteBtn.ForeColor = [System.Drawing.Color]::White
    $deleteBtn.FlatStyle = "Flat"
    $deleteBtn.Enabled = $false
    $scanButtonPanel.Controls.Add($deleteBtn)
    
    # Total size label
    $totalSizeLabel = New-Object System.Windows.Forms.Label
    $totalSizeLabel.Text = "Total selected: 0 B"
    $totalSizeLabel.Size = New-Object System.Drawing.Size(200, 25)
    $totalSizeLabel.Location = New-Object System.Drawing.Point(550, 8)
    $totalSizeLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    if ($script:DarkMode) { $totalSizeLabel.ForeColor = $script:DarkColors.TextColor }
    $scanButtonPanel.Controls.Add($totalSizeLabel)
    
    # Warning label
    $warningLabel = New-Object System.Windows.Forms.Label
    $warningLabel.Text = "WARNING: Review carefully before deleting. Some folders may contain data you want to keep!"
    $warningLabel.Size = New-Object System.Drawing.Size(760, 20)
    $warningLabel.Location = New-Object System.Drawing.Point(10, 525)
    $warningLabel.ForeColor = [System.Drawing.Color]::OrangeRed
    $scanForm.Controls.Add($warningLabel)
    
    # Folders to exclude (system/essential folders)
    $excludeFolders = @(
        "Common Files", "Windows Defender", "Windows Mail", "Windows Media Player",
        "Windows NT", "Windows Photo Viewer", "Windows Portable Devices", "Windows Security",
        "Windows Sidebar", "WindowsPowerShell", "Microsoft", "Internet Explorer",
        "Windows", "WindowsApps", "ModifiableWindowsApps", "Reference Assemblies",
        "Microsoft.NET", "dotnet", "MSBuild", "IIS", "IIS Express",
        "Microsoft Visual Studio", "Microsoft SDKs", "Microsoft SQL Server",
        "PowerShell", "NVIDIA Corporation", "AMD", "Intel", "Realtek",
        "Dell", "HP", "Lenovo", "ASUS", "Acer", "Microsoft Update Health Tools",
        "Microsoft Office", "Office", "Teams", "OneDrive", "Edge", "Windows Defender Advanced Threat Protection"
    )
    
    # Scan button click
    $scanBtn.Add_Click({
        $scanBtn.Enabled = $false
        $leftoversList.Items.Clear()
        $scanStatus.Text = "Scanning... Getting installed programs..."
        $scanProgress.Style = "Marquee"
        [System.Windows.Forms.Application]::DoEvents()
        
        # Get installed programs
        $installedPrograms = Get-InstalledPrograms
        $installedNames = $installedPrograms | ForEach-Object { $_.Name.ToLower() }
        $installedPublishers = $installedPrograms | ForEach-Object { if ($_.Publisher) { $_.Publisher.ToLower() } } | Where-Object { $_ }
        
        $leftovers = @()
        $foldersToScan = @(
            @{ Path = "C:\Program Files"; Type = "Program Files" },
            @{ Path = "C:\Program Files (x86)"; Type = "Program Files (x86)" },
            @{ Path = "$env:APPDATA"; Type = "AppData\Roaming" },
            @{ Path = "$env:LOCALAPPDATA"; Type = "AppData\Local" },
            @{ Path = "C:\ProgramData"; Type = "ProgramData" }
        )
        
        $scanProgress.Style = "Continuous"
        $scanProgress.Maximum = $foldersToScan.Count
        $scanProgress.Value = 0
        
        foreach ($scanFolder in $foldersToScan) {
            $scanProgress.Value++
            $scanStatus.Text = "Scanning $($scanFolder.Type)..."
            [System.Windows.Forms.Application]::DoEvents()
            
            if (Test-Path $scanFolder.Path) {
                try {
                    $subFolders = Get-ChildItem -Path $scanFolder.Path -Directory -ErrorAction SilentlyContinue
                    
                    foreach ($folder in $subFolders) {
                        $folderName = $folder.Name
                        $folderLower = $folderName.ToLower()
                        
                        # Skip excluded folders
                        $isExcluded = $false
                        foreach ($exclude in $excludeFolders) {
                            if ($folderLower -eq $exclude.ToLower() -or $folderLower -like "*$($exclude.ToLower())*") {
                                $isExcluded = $true
                                break
                            }
                        }
                        if ($isExcluded) { continue }
                        
                        # Check if folder matches any installed program
                        $isInstalled = $false
                        foreach ($progName in $installedNames) {
                            if ($progName -and ($folderLower -like "*$progName*" -or $progName -like "*$folderLower*")) {
                                $isInstalled = $true
                                break
                            }
                        }
                        
                        # Also check publisher names
                        if (-not $isInstalled) {
                            foreach ($publisher in $installedPublishers) {
                                if ($publisher -and $folderLower -like "*$publisher*") {
                                    $isInstalled = $true
                                    break
                                }
                            }
                        }
                        
                        # If not matched to any installed program, it might be a leftover
                        if (-not $isInstalled) {
                            # Calculate folder size
                            try {
                                $size = (Get-ChildItem -Path $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue | 
                                    Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                if (-not $size) { $size = 0 }
                            }
                            catch { $size = 0 }
                            
                            # Only add if folder has some content
                            if ($size -gt 0) {
                                $leftovers += @{
                                    Name = $folderName
                                    Path = $folder.FullName
                                    Size = $size
                                    Type = $scanFolder.Type
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }
        
        # Sort by size descending
        $leftovers = $leftovers | Sort-Object { $_.Size } -Descending
        
        # Add to list
        foreach ($leftover in $leftovers) {
            $item = New-Object System.Windows.Forms.ListViewItem($leftover.Name)
            $item.SubItems.Add($leftover.Type) | Out-Null
            $item.SubItems.Add((Format-FileSize $leftover.Size)) | Out-Null
            $item.SubItems.Add("Folder") | Out-Null
            $item.Tag = $leftover
            $leftoversList.Items.Add($item) | Out-Null
        }
        
        $scanProgress.Value = $scanProgress.Maximum
        $scanStatus.Text = "Found $($leftovers.Count) potential leftover folder(s). Review and select items to delete."
        $scanBtn.Enabled = $true
        $deleteBtn.Enabled = $leftovers.Count -gt 0
    })
    
    # Select All button click
    $selectAllBtn.Add_Click({
        $allChecked = $true
        foreach ($item in $leftoversList.Items) {
            if (-not $item.Checked) { $allChecked = $false; break }
        }
        foreach ($item in $leftoversList.Items) {
            $item.Checked = -not $allChecked
        }
        $selectAllBtn.Text = if (-not $allChecked) { "Deselect All" } else { "Select All" }
    })
    
    # Update total size when items are checked/unchecked
    $leftoversList.Add_ItemChecked({
        $totalSize = 0
        $selectedCount = 0
        foreach ($item in $leftoversList.Items) {
            if ($item.Checked -and $item.Tag) {
                $totalSize += $item.Tag.Size
                $selectedCount++
            }
        }
        $totalSizeLabel.Text = "Selected: $selectedCount ($(Format-FileSize $totalSize))"
    })
    
    # Delete button click
    $deleteBtn.Add_Click({
        $selectedItems = @()
        foreach ($item in $leftoversList.Items) {
            if ($item.Checked) {
                $selectedItems += $item
            }
        }
        
        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "No items selected for deletion.",
                "No Selection",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }
        
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to delete $($selectedItems.Count) selected folder(s)?`n`nThis action cannot be undone!`n`nTip: Some folders may contain user data or settings you want to keep.",
            "Confirm Deletion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $deleted = 0
            $failed = 0
            
            foreach ($item in $selectedItems) {
                $path = $item.Tag.Path
                try {
                    Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                    $leftoversList.Items.Remove($item)
                    $deleted++
                }
                catch {
                    $item.ForeColor = [System.Drawing.Color]::Red
                    $failed++
                }
            }
            
            $scanStatus.Text = "Deleted $deleted folder(s). $failed failed (may be in use)."
            
            if ($deleted -gt 0) {
                # Refresh disk space on main form
                Update-DiskSpace
            }
        }
    })
    
    [void]$scanForm.ShowDialog()
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
        
        # Refresh winget sources first
        Write-Log "Refreshing winget sources..."
        Update-Status "Refreshing package sources..."
        $null = winget source update 2>&1
        
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
                                Name             = $appName
                                Id               = $appId
                                CurrentVersion   = $currentVer
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

# List of package IDs that cannot be updated via winget (self-updating or Store-only)
$script:SkipPackages = @(
    "Microsoft.DesktopAppInstaller",
    "Microsoft.WindowsStore",
    "Microsoft.StorePurchaseApp",
    "Microsoft.WindowsTerminal",
    "Microsoft.GetHelp",
    "Microsoft.Getstarted"
)

# Mapping of package IDs to process names that need to be closed before updating
$script:ProcessesToClose = @{
    "Mozilla.Firefox"           = @("firefox", "plugin-container", "crashreporter", "updater", "pingsender", "default-browser-agent", "maintenanceservice")
    "Mozilla.Thunderbird"       = @("thunderbird", "plugin-container", "crashreporter", "updater")
    "Google.Chrome"             = @("chrome", "GoogleUpdate", "crashpad_handler")
    "Microsoft.Edge"            = @("msedge", "MicrosoftEdgeUpdate")
    "Opera.Opera"               = @("opera", "opera_crashreporter", "launcher")
    "BraveSoftware.BraveBrowser"= @("brave", "BraveUpdate")
    "Vivaldi.Vivaldi"           = @("vivaldi")
    "Discord.Discord"           = @("discord", "update")
    "Spotify.Spotify"           = @("spotify")
    "SlackTechnologies.Slack"   = @("slack")
    "Microsoft.Teams"           = @("teams", "ms-teams")
    "Zoom.Zoom"                 = @("zoom")
    "Valve.Steam"               = @("steam", "steamwebhelper", "steamservice")
    "EpicGames.EpicGamesLauncher" = @("epicgameslauncher")
    "Notepad++.Notepad++"       = @("notepad++")
    "VideoLAN.VLC"              = @("vlc")
}

# Services that should be stopped before updating certain packages
$script:ServicesToStop = @{
    "Mozilla.Firefox"    = @("MozillaMaintenance")
    "Mozilla.Thunderbird"= @("MozillaMaintenance")
    "Google.Chrome"      = @("gupdate", "gupdatem")
    "Microsoft.Edge"     = @("edgeupdate", "edgeupdatem")
}

# Function to close processes for a package
function Close-ProcessesForPackage {
    param([string]$PackageId)
    
    $closedAny = $false
    
    # Stop related services first
    foreach ($key in $script:ServicesToStop.Keys) {
        if ($PackageId -like "*$key*" -or $key -like "*$PackageId*") {
            $serviceNames = $script:ServicesToStop[$key]
            foreach ($svcName in $serviceNames) {
                $svc = Get-Service -Name $svcName -ErrorAction SilentlyContinue
                if ($svc -and $svc.Status -eq 'Running') {
                    Write-Log "Stopping service $svcName before update..."
                    try {
                        Stop-Service -Name $svcName -Force -ErrorAction SilentlyContinue
                        $closedAny = $true
                    }
                    catch { }
                }
            }
        }
    }
    
    # Then close processes
    foreach ($key in $script:ProcessesToClose.Keys) {
        if ($PackageId -like "*$key*" -or $key -like "*$PackageId*") {
            $processNames = $script:ProcessesToClose[$key]
            foreach ($procName in $processNames) {
                $procs = Get-Process -Name $procName -ErrorAction SilentlyContinue
                if ($procs) {
                    Write-Log "Closing $procName before update..."
                    $procs | ForEach-Object { 
                        try {
                            $_.CloseMainWindow() | Out-Null
                            Start-Sleep -Milliseconds 500
                            if (!$_.HasExited) {
                                $_ | Stop-Process -Force -ErrorAction SilentlyContinue
                            }
                        }
                        catch { }
                    }
                    $closedAny = $true
                }
            }
        }
    }
    
    if ($closedAny) {
        Start-Sleep -Seconds 3  # Give processes and services time to fully close
    }
    
    return $closedAny
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
        
        # Skip packages that can't be updated via winget
        if ($script:SkipPackages -contains $appInfo.Id) {
            $item.SubItems[3].Text = "Skipped"
            Write-Log "Skipped: $($appInfo.Name) (requires Microsoft Store)"
            $progressBar.Value = [math]::Min(100, [int](($current / $selectedApps.Count) * 100))
            [System.Windows.Forms.Application]::DoEvents()
            continue
        }
        
        Update-Status "Updating: $($appInfo.Name) ($current/$($selectedApps.Count))..."
        $item.SubItems[3].Text = "Installing..."
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            # Close any running processes that might block the update
            $closedProcesses = Close-ProcessesForPackage -PackageId $appInfo.Id
            if ($closedProcesses) {
                Write-Log "Closed running processes for $($appInfo.Name)"
            }
            
            # Some packages need --silent to avoid requiring user interaction
            # Firefox, Chrome, etc. support silent installs
            $silentPackages = @("Mozilla.Firefox", "Mozilla.Thunderbird", "Google.Chrome", "VideoLAN.VLC", "Notepad++.Notepad++")
            $useSilent = $false
            foreach ($pkg in $silentPackages) {
                if ($appInfo.Id -like "*$pkg*") {
                    $useSilent = $true
                    break
                }
            }
            
            if ($useSilent) {
                $result = winget upgrade --id $appInfo.Id --silent --force --accept-package-agreements --accept-source-agreements --disable-interactivity 2>&1 | Out-String
            }
            else {
                $result = winget upgrade --id $appInfo.Id --force --accept-package-agreements --accept-source-agreements --disable-interactivity 2>&1 | Out-String
            }
            
            if ($LASTEXITCODE -eq 0) {
                $item.SubItems[3].Text = "Updated"
                Write-Log "Updated: $($appInfo.Name)"
            }
            else {
                $item.SubItems[3].Text = "Failed"
                # Extract meaningful error from output
                $errorMsg = if ($result -match "0x[0-9a-fA-F]+") { $Matches[0] } else { "Exit code: $LASTEXITCODE" }
                Write-Log "Failed to update: $($appInfo.Name) ($errorMsg)"
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
        # Check network first
        if (-not (Test-NetworkConnection)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No internet connection detected.`n`nPlease check your network connection and try again.",
                "Network Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }
        
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
        Update-DiskSpace
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
        Update-DiskSpace
    
        # Show reboot prompt if needed
        if ($rebootRequired) {
            $rebootButton.Visible = $true
            $rebootAutoButton.Visible = $true
            $cleanButton.Visible = $false
            Update-Status "Updates completed. A reboot is required to finish installation."
        
            if ($autoRestartCheck.Checked) {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Updates have been installed successfully.`n`nA system restart is required to complete the installation.`n`nWould you like to restart now?`n`nThe updater will relaunch automatically after restart.",
                    "Restart Required",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
            
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-Log "User initiated system restart with auto-relaunch"
                    Register-AutoRestart
                    Restart-Computer -Force
                }
            }
            else {
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
        }
        else {
            $result = [System.Windows.Forms.MessageBox]::Show(
                "All updates have been installed successfully!`n`nIt's recommended to restart your computer.`n`nWould you like to restart now?",
                "Updates Complete",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                if ($autoRestartCheck.Checked) {
                    Register-AutoRestart
                }
                Write-Log "User initiated system restart"
                Restart-Computer -Force
            }
            else {
                $rebootButton.Visible = $true
                $rebootAutoButton.Visible = $true
                $cleanButton.Visible = $false
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

# Reboot with Auto-Start Button Click Event
$rebootAutoButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to restart your computer now?`n`nThe updater will relaunch automatically after restart.`n`nPlease save all your work before continuing.",
            "Confirm Restart with Auto-Relaunch",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Write-Log "Initiating system restart with auto-relaunch..."
            Register-AutoRestart
            Restart-Computer -Force
        }
    })

# Dark Mode Button Click Event
$darkModeButton.Add_Click({
        Set-DarkMode (-not $script:DarkMode)
    })

# History Button Click Event
$historyButton.Add_Click({
        Show-UpdateHistory
    })

# Leftovers Button Click Event
$leftoversButton.Add_Click({
        Show-LeftoverScanner
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
        
            Update-DiskSpace
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

# Initialize on startup
Update-DiskSpace
Test-NetworkConnection

# Check if this is an auto-restart run
if ($script:IsAutoRestart) {
    Write-Log "Resumed after system restart"
    Update-Status "Welcome back! Click 'Check for Updates' to continue."
    [System.Windows.Forms.MessageBox]::Show(
        "The system has restarted successfully.`n`nYou can now check for additional updates or clean temporary files.",
        "Restart Complete",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

Write-Log "Windows System Updater started"

# Show the form
[void]$form.ShowDialog()
