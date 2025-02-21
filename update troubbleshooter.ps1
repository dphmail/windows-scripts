# Windows Update Troubleshooting Script
# This script does NOT reboot the server or require a reboot afterward.

Write-Host "Starting Windows Update Troubleshooting..." -ForegroundColor Cyan

# 1. Check Windows Update Service Status
Write-Host "`nChecking Windows Update Service Status..." -ForegroundColor Yellow
$wuauserv = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($wuauserv.Status -eq "Running") {
    Write-Host "Windows Update Service is running." -ForegroundColor Green
} else {
    Write-Host "Windows Update Service is NOT running. Attempting to start..." -ForegroundColor Red
    Start-Service wuauserv -ErrorAction SilentlyContinue
    if ($?) {
        Write-Host "Service started successfully." -ForegroundColor Green
    } else {
        Write-Host "Failed to start service." -ForegroundColor Red
    }
}

# 2. Check Windows Update Log for Recent Errors
Write-Host "`nChecking Windows Update Log for Errors..." -ForegroundColor Yellow
$updateErrors = Get-WinEvent -LogName "Microsoft-Windows-WindowsUpdateClient/Operational" -MaxEvents 10 |
    Where-Object { $_.Level -eq 2 }
if ($updateErrors) {
    Write-Host "Recent Windows Update Errors Found:" -ForegroundColor Red
    $updateErrors | Format-Table -AutoSize
} else {
    Write-Host "No critical Windows Update errors found in the logs." -ForegroundColor Green
}

# 3. Check Disk Space on C: Drive
Write-Host "`nChecking Disk Space on C: Drive..." -ForegroundColor Yellow
$disk = Get-PSDrive -Name C
if ($disk.Free -lt 10GB) {
    Write-Host "Warning: Low Disk Space on C: Drive! Only $([math]::Round($disk.Free / 1GB,2)) GB left. Proceeding to attempt to clear space" -ForegroundColor Red
    
    # Disk cleaning steps with progress bar
    $progress = 0
    Write-Progress -Activity "Clearing Disk Space" -Status "Starting DISM cleanup" -PercentComplete $progress

    # Run DISM cleanup and wait for completion
    Start-Process -FilePath "dism" -ArgumentList "/online /cleanup-image /spsuperseded" -NoNewWindow -Wait
    $progress += 25
    Write-Progress -Activity "Clearing Disk Space" -Status "DISM cleanup completed" -PercentComplete $progress

    # Run Cleanmgr with /verylowdisk and wait for completion
    Write-Progress -Activity "Clearing Disk Space" -Status "Running Cleanmgr (/verylowdisk)" -PercentComplete $progress
    Start-Process -FilePath "cleanmgr" -ArgumentList "/verylowdisk" -NoNewWindow -Wait
    $progress += 25
    Write-Progress -Activity "Clearing Disk Space" -Status "Cleanmgr (/verylowdisk) completed" -PercentComplete $progress

    # Run Cleanmgr with /autoclean and wait for completion
    Write-Progress -Activity "Clearing Disk Space" -Status "Running Cleanmgr (/autoclean)" -PercentComplete $progress
    Start-Process -FilePath "cleanmgr" -ArgumentList "/autoclean" -NoNewWindow -Wait
    $progress += 25
    Write-Progress -Activity "Clearing Disk Space" -Status "Cleanmgr (/autoclean) completed" -PercentComplete $progress

    # Remove temporary directories
    if (Test-Path "$env:SystemRoot\Installer\PatchCache") {
        Write-Progress -Activity "Clearing Disk Space" -Status "Removing PatchCache" -PercentComplete $progress
        Remove-Item -Recurse -Force "$env:SystemRoot\Installer\PatchCache"
        $progress += 5
    }
    if (Test-Path "$env:ALLUSERSPROFILE\Microsoft\Windows\WER") {
        Write-Progress -Activity "Clearing Disk Space" -Status "Removing WER logs" -PercentComplete $progress
        Remove-Item -Recurse -Force "$env:ALLUSERSPROFILE\Microsoft\Windows\WER"
        $progress += 5
    }
    
    # Stop Windows Update service, clear SoftwareDistribution, then restart service
    Write-Progress -Activity "Clearing Disk Space" -Status "Stopping Windows Update service" -PercentComplete $progress
    Stop-Service wuauserv
    if (Test-Path "$env:SystemRoot\SoftwareDistribution") {
        Write-Progress -Activity "Clearing Disk Space" -Status "Removing SoftwareDistribution folder" -PercentComplete $progress
        Remove-Item -Recurse -Force "$env:SystemRoot\SoftwareDistribution"
        $progress += 5
    }
    Write-Progress -Activity "Clearing Disk Space" -Status "Restarting Windows Update service" -PercentComplete $progress
    Start-Service wuauserv

    Write-Progress -Activity "Clearing Disk Space" -Status "Disk cleanup complete" -PercentComplete 100 -Completed

    Write-Host "`nDisk space cleared, re-checking Disk Space on C: Drive..." -ForegroundColor Yellow
    $disk = Get-PSDrive -Name C
    if ($disk.Free -lt 10GB) {
        Write-Host "Warning: Did not manage to clear enough space. Only $([math]::Round($disk.Free / 1GB,2)) GB left." -ForegroundColor Red
    } else {
        Write-Host "Sufficient disk space available: $([math]::Round($disk.Free / 1GB,2)) GB" -ForegroundColor Green
    }
} else {
    Write-Host "Sufficient disk space available: $([math]::Round($disk.Free / 1GB,2)) GB" -ForegroundColor Green
}

# 4. Check for Pending Updates, needs more work to only get latest ones
Write-Host "`nChecking for Pending Windows Updates..." -ForegroundColor Yellow

# Create a Windows Update session and searcher
$updateSession = New-Object -ComObject Microsoft.Update.Session
$updateSearcher = $updateSession.CreateUpdateSearcher()

# Search for updates that are not installed and not hidden
$searchResult = $updateSearcher.Search("IsInstalled=0 and IsHidden=0")

# Process each update returned
$searchResult.Updates | ForEach-Object {
    $update = $_
    # Get a list of classification names from the update's categories
    $categories = $update.Categories | ForEach-Object { $_.Name }
    # If the update is classified as either Critical or Security, process it further
    if ($categories -contains "Critical Updates" -or $categories -contains "Security Updates") {
        # Combine all KB Article IDs (if more than one) into a single string
        $kbIDs = if ($update.KBArticleIDs) { $update.KBArticleIDs -join ", " } else { "N/A" }
        # Identify the matching classification(s)
        $classification = ($categories | Where-Object { $_ -eq "Critical Updates" -or $_ -eq "Security Updates" } |
            Sort-Object -Unique) -join ", "
        # Output the selected information as a custom object
        [PSCustomObject]@{
            "KB Name"        = $kbIDs
            "Classification" = $classification
            "Title"          = $update.Title
        }
    }
} | Format-Table -AutoSize

# 5. Checking if server can reach its update source

# Check if WSUS is configured by examining the registry
$wsusAU = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -ErrorAction SilentlyContinue

if ($wsusAU -and $wsusAU.UseWUServer -eq 1) {
    $wsusSettings = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -ErrorAction SilentlyContinue
    if ($wsusSettings.WUServer) {
        try {
            # Convert the WSUS server string to a URI object for parsing
            $wsusUri = [Uri]$wsusSettings.WUServer
            Write-Host "`nWSUS is configured. Checking connectivity to WSUS server at $($wsusUri.Host) on port $($wsusUri.Port) ..." -ForegroundColor Yellow
            $wsusConn = Test-NetConnection -ComputerName $wsusUri.Host -Port $wsusUri.Port
            if ($wsusConn.TcpTestSucceeded) {
                Write-Host "WSUS server is reachable." -ForegroundColor Green
                Write-Host "Warning: Updates displayed may not be accurate as only approved WSUS updates are shown." -ForegroundColor Yellow
            } else {
                Write-Host "WSUS server is NOT reachable. Check your network or firewall settings." -ForegroundColor Red
            }
        } catch {
            Write-Host "Failed to parse WSUS server settings. Falling back to Windows Update connectivity check..." -ForegroundColor Red
            $wuComponents = Test-NetConnection -ComputerName "update.microsoft.com" -Port 443
            if ($wuComponents.TcpTestSucceeded) {
                Write-Host "Windows Update servers are reachable." -ForegroundColor Green
            } else {
                Write-Host "Windows Update servers are NOT reachable. Check your network or firewall settings." -ForegroundColor Red
            }
        }
    } else {
        Write-Host "WSUS configuration detected but no WUServer value found. Checking Windows Update source..." -ForegroundColor Yellow
        $wuComponents = Test-NetConnection -ComputerName "update.microsoft.com" -Port 443
        if ($wuComponents.TcpTestSucceeded) {
            Write-Host "Windows Update servers are reachable." -ForegroundColor Green
        } else {
            Write-Host "Windows Update servers are NOT reachable. Check your network or firewall settings." -ForegroundColor Red
        }
    }
} else {
    Write-Host "`nChecking Windows Update source..." -ForegroundColor Yellow
    $wuComponents = Test-NetConnection -ComputerName "update.microsoft.com" -Port 443
    if ($wuComponents.TcpTestSucceeded) {
        Write-Host "Windows Update servers are reachable." -ForegroundColor Green
    } else {
        Write-Host "Windows Update servers are NOT reachable. Check your network or firewall settings." -ForegroundColor Red
    }
}

# 6. Run System File Checker (SFC) and DISM (Without Reboot)

# Run SFC with a simulated progress bar
Write-Host "`nRunning System File Checker (SFC) to Check for Corrupt Files..." -ForegroundColor Yellow
$jobSFC = Start-Job -ScriptBlock { sfc /scannow }
while ($jobSFC.State -eq 'Running') {
    Write-Progress -Activity "Running SFC" -Status "Scanning for corrupt files..." -PercentComplete (Get-Random -Minimum 1 -Maximum 90)
    Start-Sleep -Seconds 5
}
Receive-Job $jobSFC | Out-Null
Write-Progress -Activity "Running SFC" -Status "SFC scan completed." -PercentComplete 100 -Completed
Write-Host "SFC scan completed." -ForegroundColor Green

# Run DISM with a simulated progress bar
Write-Host "`nRunning DISM to Check Windows Update Health..." -ForegroundColor Yellow
$jobDISM = Start-Job -ScriptBlock { DISM /Online /Cleanup-Image /CheckHealth }
while ($jobDISM.State -eq 'Running') {
    Write-Progress -Activity "Running DISM" -Status "Checking Windows Update Health..." -PercentComplete (Get-Random -Minimum 1 -Maximum 90)
    Start-Sleep -Seconds 5
}
Receive-Job $jobDISM | Out-Null
Write-Progress -Activity "Running DISM" -Status "DISM check completed." -PercentComplete 100 -Completed
Write-Host "DISM check completed. If issues were found, run 'DISM /Online /Cleanup-Image /RestoreHealth' manually." -ForegroundColor Green

Write-Host "`nWindows Update Troubleshooting Completed." -ForegroundColor Cyan
