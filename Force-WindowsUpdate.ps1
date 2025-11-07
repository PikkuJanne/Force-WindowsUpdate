<#
Force-WindowsUpdate.ps1
One-click Windows Update fixer for non-technical users

Author: Janne Vuorela
Target OS: Windows 10/11
PowerShell: Windows PowerShell 5.1 (built-in)
Dependencies: PSWindowsUpdate module (from PSGallery), .cmd wrapper for one-click launch 

SYNOPSIS
    Resets common Windows Update components, runs DISM and SFC to repair system files. 
    Then uses PSWindowsUpdate to search for and install all available updates. 
    Writes a transcript to the user's Desktop and keeps the console open so they can read what happened.

WHAT THIS IS (AND ISN'T)
    - Personal, purpose-built helper for "run this first" Windows Update issues.
    - Favors predictable, repeatable behavior over configurability.
    - Designed as a safe, verbose "grandma's little updater" for family/friends.
    - Not a full Windows Update diagnostic suite and not a way around corporate policies, WSUS, or deep OS corruption.

FEATURES
    - Admin check:
        Verifies the script is running as administrator. 
        If not, prints clear right-click / "Run as administrator" instructions and exits cleanly.
    - Component reset:
        Stops core update services (wuauserv, bits, cryptsvc, msiserver).
        Renames SoftwareDistribution and catroot2 to timestamped .bak folders,
        Then starts the services again. All wrapped in try/catch with gentle "not always a problem" wording.
    - System file repair:
        Runs:
            dism.exe /Online /Cleanup-Image /RestoreHealth
            sfc.exe /scannow
        Explains that these can take a long time and that slowdowns are normal, and prints friendly summaries for common exit codes.
    - Windows Update via PSWindowsUpdate:
        Ensures TLS 1.2 is enabled where needed.
        Tries to install NuGet and PSWindowsUpdate from PSGallery, imports the module, then runs:
        Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Verbose
        to pull and install available updates.
    - Logging:
        Creates Desktop\WindowsUpdateHelper_Logs\
        Starts a transcript:
            WindowsUpdateHelper_yyyyMMdd_HHmmss.txt
        You can review or ask users to send you the latest log.
    - User-friendly output:
        Very verbose step-by-step console messages, 5 clearly labeled steps.
        Plain-language explanations aimed at non-technical users.
    - Safe exit:
        Uses try/catch/finally and a Read-Host prompt. 
        Window never closes automatically, it always waits for ENTER even when errors occur.

MY INTENDED USAGE
    - I share the .ps1 and .cmd files with family/friends.
    - I tell them: 
        Right-click Run-WindowsUpdateHelper.cmd, 
        choose Run as administrator, 
        let it finish, 
        then send me the newest WindowsUpdateHelper_*.txt from your Desktop if it still fails.

SETUP
    1) Copy these files into a directory, for example:
        C:\Tools\GrandmasLittleUpdater\
           - Force-WindowsUpdate.ps1
           - Run-WindowsUpdateHelper.cmd
    2) (Optional) Create a Desktop shortcut to Run-WindowsUpdateHelper.cmd, rename it to something friendly like "Windows Update Helper".
    3) On first use, make sure you can run the .cmd as administrator on the target machines (UAC, AV, and policies permitting).

USAGE
    A) Right-click .cmd launcher 
       - Right-click Run-WindowsUpdateHelper.cmd
       - Click "Run as administrator"
       - Accept the UAC prompt
       - Read the messages, wait for "ALL DONE", then press ENTER to close

    B) Direct PowerShell 
       - Open PowerShell as administrator
       - cd C:\Tools\GrandmasLittleUpdater\
       - .\Force-WindowsUpdate.ps1

NOTES
    - The script renames SoftwareDistribution and catroot2 to timestamped .bak_yyyyMMddHHmmss folders instead of deleting them outright.
    - It does not create scheduled tasks, registry keys, or permanent changes beyond the standard DISM/SFC repairs and update cache reset.
    - Designed primarily for home PCs, corporate or domain-joined devices may have policies that limit what it can do.

LIMITATIONS
    - Cannot override WSUS, Group Policy, or other enterprise update settings.
    - DISM and SFC can only repair what is still repairable, heavily damaged systems may still need in-place upgrade or reinstall.
    - PSWindowsUpdate install/import may fail on machines without internet or where PSGallery is blocked.
    - Third-party AV or endpoint protection can still interfere with updates, this script does not attempt to disable or modify security products.

TROUBLESHOOTING
    - "This tool must be run as administrator":
        User launched it without elevation. Have them right-click the .cmd and choose "Run as administrator", then accept UAC.
    - Stuck on DISM or SFC:
        Normal on slow HDD-based systems. As long as the percentage changes occasionally, let it run. 
        If truly stuck for a very long time, disk health may be an issue.
    - PSWindowsUpdate cannot be installed/imported:
        Usually no internet or PowerShell Gallery is blocked by policy. The script will say so and ask the user to send you the transcript.
    - Updates still fail after the script:
        Ask for:
            - The latest WindowsUpdateHelper_*.txt from the Desktop.
            - A screenshot of the Windows Update error.
        Then follow up with more targeted tools or manual steps.

LICENSE / WARRANTY
    - Personal tool, provided as-is without warranty. Use at your own risk.
    - Feel free to fork, trim, or extend it for your own support workflows.
#>

param()

$ErrorActionPreference = 'Stop'

$desktop = [Environment]::GetFolderPath('Desktop')
$logDir  = Join-Path $desktop "WindowsUpdateHelper_Logs"
$logPath = $null

try {
    # Try to set a clear window title, ignore errors
    try {
        $Host.UI.RawUI.WindowTitle = "Windows Update Helper - Please do NOT close this window"
    } catch {
    }

    # --- 1. Prepare logging ---
    New-Item -Path $logDir -ItemType Directory -Force -ErrorAction SilentlyContinue | Out-Null

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logPath   = Join-Path $logDir "WindowsUpdateHelper_$timestamp.txt"

    try {
        Start-Transcript -Path $logPath -Append -ErrorAction SilentlyContinue | Out-Null
    } catch {
    }

    # --- Intro text ---
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "             WINDOWS UPDATE HELPER"               -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "This tool will:" -ForegroundColor Yellow
    Write-Host "  1) Try to fix common Windows Update problems." -ForegroundColor Yellow
    Write-Host "  2) Check and repair important system files."  -ForegroundColor Yellow
    Write-Host "  3) Search for updates and install them."      -ForegroundColor Yellow
    Write-Host ""
    Write-Host "While it works:" -ForegroundColor Yellow
    Write-Host "  - Do NOT close this window."                  -ForegroundColor Yellow
    Write-Host "  - Your PC may feel slower and fans may spin." -ForegroundColor Yellow
    Write-Host ""

    if ($logDir -and $logPath) {
        Write-Host "Everything printed here is also saved to a file on your Desktop:" -ForegroundColor Yellow
        Write-Host "  $logPath" -ForegroundColor White
        Write-Host ""
    }

    # --- 2. Admin check ---
    $isAdmin = $false
    try {
        $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal       = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
        $isAdmin         = $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    } catch {
        $isAdmin = $false
    }

    if (-not $isAdmin) {
        Write-Host "IMPORTANT:" -ForegroundColor Red
        Write-Host "This tool must be run as administrator to work properly." -ForegroundColor Red
        Write-Host ""
        Write-Host "Please do this:" -ForegroundColor Yellow
        Write-Host "  1) Close this window." -ForegroundColor Yellow
        Write-Host "  2) Right-click the .cmd or .ps1 file you got." -ForegroundColor Yellow
        Write-Host "  3) Choose: 'Run as administrator' or 'Run with PowerShell'." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "If you do not see these options, ask your helper for instructions." -ForegroundColor Yellow
        return
    }

    # --- 3. Basic system info ---
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "STEP 1/5 - Checking basic information"           -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan

    try {
        $os = Get-CimInstance Win32_OperatingSystem
        Write-Host "Computer name : $env:COMPUTERNAME"
        Write-Host "Logged in user: $env:USERNAME"
        Write-Host "Windows       : $($os.Caption) ($($os.Version))"
    } catch {
        Write-Host "Could not read detailed Windows information (this is not critical)." -ForegroundColor DarkYellow
    }
    Write-Host ""

    # --- 4. Reset Windows Update components ---
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "STEP 2/5 - Resetting Windows Update components"  -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "Stopping Windows Update related services..."     -ForegroundColor Yellow

    $services = 'wuauserv','bits','cryptsvc','msiserver'

    foreach ($svc in $services) {
        try {
            Write-Host "  - Stopping service: $svc" -NoNewline
            Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
            Write-Host "  ...OK"
        } catch {
            Write-Host "  ...could not stop (this is not always a problem)" -ForegroundColor DarkYellow
        }
    }

    Write-Host ""
    Write-Host "Cleaning Windows Update cache folders..." -ForegroundColor Yellow

    $windowsFolder = $env:windir
    $timeTag       = Get-Date -Format 'yyyyMMddHHmmss'

    $sdPath   = Join-Path $windowsFolder "SoftwareDistribution"
    $sdBackup = $sdPath + ".bak_$timeTag"

    if (Test-Path $sdPath) {
        Write-Host "  - Renaming: $sdPath" -ForegroundColor White
        try {
            Rename-Item -Path $sdPath -NewName $sdBackup -ErrorAction SilentlyContinue
            Write-Host "    Saved old folder as: $sdBackup" -ForegroundColor DarkGray
        } catch {
            Write-Host "    Could not rename folder (might be locked). Skipping." -ForegroundColor DarkYellow
        }
    } else {
        Write-Host "  - SoftwareDistribution folder not found (already cleaned?)." -ForegroundColor DarkYellow
    }

    $catrootPath   = Join-Path (Join-Path $windowsFolder "System32") "catroot2"
    $catrootBackup = $catrootPath + ".bak_$timeTag"

    if (Test-Path $catrootPath) {
        Write-Host "  - Renaming: $catrootPath" -ForegroundColor White
        try {
            Rename-Item -Path $catrootPath -NewName $catrootBackup -ErrorAction SilentlyContinue
            Write-Host "    Saved old folder as: $catrootBackup" -ForegroundColor DarkGray
        } catch {
            Write-Host "    Could not rename folder (might be locked). Skipping." -ForegroundColor DarkYellow
        }
    } else {
        Write-Host "  - catroot2 folder not found (already cleaned?)." -ForegroundColor DarkYellow
    }

    Write-Host ""
    Write-Host "Starting Windows Update related services again..." -ForegroundColor Yellow

    foreach ($svc in $services) {
        try {
            Write-Host "  - Starting service: $svc" -NoNewline
            Start-Service -Name $svc -ErrorAction SilentlyContinue
            Write-Host "  ...OK"
        } catch {
            Write-Host "  ...could not start - Windows may start it later automatically." -ForegroundColor DarkYellow
        }
    }

    # --- 5. DISM and SFC ---
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "STEP 3/5 - Checking and repairing system files"  -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan

    Write-Host "First check: DISM (deep health check)." -ForegroundColor Yellow
    Write-Host "This can take 10 to 30 minutes. Please be patient." -ForegroundColor Yellow
    Write-Host "Do NOT close this window while progress is running." -ForegroundColor Yellow
    Write-Host ""

    $dismArgs = "/Online","/Cleanup-Image","/RestoreHealth"
    $dism     = Start-Process -FilePath "dism.exe" -ArgumentList $dismArgs -Wait -PassThru

    if ($dism.ExitCode -eq 0) {
        Write-Host "DISM finished successfully." -ForegroundColor Green
    } else {
        Write-Host "DISM finished with error code $($dism.ExitCode)." -ForegroundColor Red
        Write-Host "This does not always mean Windows cannot update," -ForegroundColor Yellow
        Write-Host "but your helper may want to read the log file later." -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host "Second check: SFC (System File Checker)." -ForegroundColor Yellow
    Write-Host "This can also take several minutes."      -ForegroundColor Yellow
    Write-Host ""

    $sfc = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -PassThru

    switch ($sfc.ExitCode) {
        0 { Write-Host "SFC did not find integrity violations."             -ForegroundColor Green }
        1 { Write-Host "SFC found problems and fixed them."                 -ForegroundColor Yellow }
        default {
            Write-Host "SFC finished with exit code $($sfc.ExitCode)."      -ForegroundColor Red
            Write-Host "Some files might not have been repaired completely." -ForegroundColor Yellow
        }
    }

    # --- 6. Prepare PSWindowsUpdate module ---
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "STEP 4/5 - Preparing Windows Update tools"        -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan

    Write-Host "Getting the helper tools needed to talk to Windows Update..." -ForegroundColor Yellow

    $moduleName = "PSWindowsUpdate"

    try {
        if (-not (Get-Module -ListAvailable -Name $moduleName)) {
            Write-Host "Installing PowerShell module '$moduleName' from the gallery..." -ForegroundColor Yellow

            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            } catch {
            }

            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue

            Install-Module -Name $moduleName -Force -ErrorAction Stop
            Write-Host "Module '$moduleName' installed." -ForegroundColor Green
        } else {
            Write-Host "Module '$moduleName' is already installed." -ForegroundColor Green
        }

        Import-Module -Name $moduleName -Force -ErrorAction Stop
        Write-Host "Module '$moduleName' loaded successfully." -ForegroundColor Green
    }
    catch {
        Write-Host ""
        Write-Host "Could not prepare the helper tools." -ForegroundColor Red
        Write-Host "This often happens on work computers managed by an IT department" -ForegroundColor Yellow
        Write-Host "or if there is no Internet connection."                           -ForegroundColor Yellow
        Write-Host ""
        if ($logPath) {
            Write-Host "Please send this log file to your helper:" -ForegroundColor Yellow
            Write-Host "  $logPath" -ForegroundColor White
        }
        return
    }

    # --- 7. Search and install updates ---
    Write-Host ""
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan
    Write-Host "STEP 5/5 - Searching for and installing updates"  -ForegroundColor Cyan
    Write-Host "--------------------------------------------------" -ForegroundColor Cyan

    Write-Host "Now asking Windows Update for new updates." -ForegroundColor Yellow
    Write-Host "This can again take quite a while. Please wait..." -ForegroundColor Yellow
    Write-Host ""

    try {
        $updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Verbose -ErrorAction Stop

        if (-not $updates) {
            Write-Host ""
            Write-Host "Windows did not report any updates to install." -ForegroundColor Green
            Write-Host "Your PC looks up to date right now."            -ForegroundColor Green
        } else {
            $count = $updates.Count
            Write-Host ""
            Write-Host "$count update(s) were processed." -ForegroundColor Green
            Write-Host "Scroll up in this window to see:" -ForegroundColor Green
            Write-Host "  - Which updates were installed." -ForegroundColor Green
            Write-Host "  - Whether a restart is required." -ForegroundColor Green
        }
    }
    catch {
        Write-Host ""
        Write-Host "Something went wrong while talking to Windows Update." -ForegroundColor Red
        Write-Host "Sometimes simply restarting the PC and running this tool again helps." -ForegroundColor Yellow
        Write-Host ""
        if ($logPath) {
            Write-Host "If the problem keeps coming back, send this log file to your helper:" -ForegroundColor Yellow
            Write-Host "  $logPath" -ForegroundColor White
        }
    }

    # --- 8. Final messages ---
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "                     ALL DONE                     " -ForegroundColor Cyan
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next steps for you:" -ForegroundColor Yellow
    Write-Host "  1) If you see 'Restart required' above, please save your work and restart the computer." -ForegroundColor Yellow
    Write-Host "  2) After restart, open Settings -> Windows Update and check if it still shows any errors." -ForegroundColor Yellow

    if ($logPath) {
        Write-Host "  3) If problems remain, send this file to your helper:" -ForegroundColor Yellow
        Write-Host "       $logPath" -ForegroundColor White
    }

}
catch {
    Write-Host ""
    Write-Host "UNEXPECTED ERROR OCCURRED:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    if ($logPath) {
        Write-Host "Please send this log file to your helper so they can see what went wrong:" -ForegroundColor Yellow
        Write-Host "  $logPath" -ForegroundColor White
    }
}
finally {
    try {
        Stop-Transcript | Out-Null
    } catch {
    }

    Write-Host ""
    [void](Read-Host "Press ENTER to close this window")
}
