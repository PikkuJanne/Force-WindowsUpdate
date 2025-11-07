# Force-WindowsUpdate — One-click Windows Update fixer (PowerShell)

A small, overprotective helper for family and friends whose Windows Updates are “always broken”.  
This Grandma’s Little Updater tries the usual deep-clean steps for Windows Update, repairs system files, and then asks Windows to pull in all available updates. It’s intentionally narrow and boring, one job, done the same way every time. I use it as a “please run this first” tool before remoting in. The goal is to cut down on remote sessions and give non-technical users something safe they can one-click and watch.

**Synopsis**

- Resets common Windows Update components (services + cache folders).
- Runs a deep health check and system file repair (DISM + SFC).
- Installs/loads the PSWindowsUpdate module (if allowed) and runs Get-WindowsUpdate with:
  - -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot
- Writes a full transcript to the user’s Desktop (timestamped log folder).
- Very verbose, step-by-step console text aimed at non-technical users.
- Window always stays open and asks for ENTER at the end (even on error).

**Requirements**

- Windows 10 or 11  
- Windows PowerShell 5.1 (built-in)  
- Local admin rights on the machine  
- Working internet connection for actual updates (and for installing PSWindowsUpdate)  

**Nice to have**

- A bit of free disk space: the script renames SoftwareDistribution and catroot2 to timestamped .bak_YYYYMMDDHHMMSS folders instead of deleting them.

**Files**

Place these together (e.g. C:\Tools\GrandmasLittleUpdater\):
- Force-WindowsUpdate.ps1 
  - Main script: resets components, runs DISM+SFC, installs PSWindowsUpdate, installs updates, logs everything.
- Run-WindowsUpdateHelper.cmd  
  Simple launcher: users right-click this and choose **Run as administrator**

**Installation**

Copy the two files to a folder of your choice, e.g.:
C:\Tools\GrandmasLittleUpdater\
(Optional, but recommended for family use)
Create a desktop shortcut to Run-WindowsUpdateHelper.cmd and name it something friendly:
“Windows Update Helper”
or “Grandmas Little Updater”

Tell people:
Right-click this icon and choose "Run as administrator".
That’s it. No global install, no registry changes, no scheduled tasks.

**Usage**

Recommended: Right-click the .cmd launcher

For non-technical users:
Right-click Run-WindowsUpdateHelper.cmd.
Click Run as administrator.
If Windows asks “Do you want to allow this app to make changes”, click Yes.
Read the messages on screen and wait until it says ALL DONE and asks to press ENTER.

The console shows:
A short explanation of what the tool does.
5 numbered steps:
Step 1/5 – basic system info
Step 2/5 – reset Windows Update components
Step 3/5 – DISM + SFC
Step 4/5 – prepare update tools
Step 5/5 – search for and install updates
Clear “what to do next” instructions (restart if needed, check Windows Update, send log if problems remain).
The window never closes automatically – it always waits for ENTER.

**Command line**

Run from an elevated PowerShell prompt:

cd C:\Tools\GrandmasLittleUpdater\
.\Force-WindowsUpdate.ps1

You’ll see the same step-by-step output and the same final “Press ENTER to close this window” prompt.

**What it actually does (step-by-step)**
Admin check
Verifies the script is running as administrator.
If not, prints simple instructions (“right-click, run as administrator”) and exits cleanly.

Log setup
Creates WindowsUpdateHelper_Logs on the current user’s Desktop.
Starts a PowerShell transcript to:
Desktop\WindowsUpdateHelper_Logs\WindowsUpdateHelper_yyyyMMdd_HHmmss.txt

Basic info
Prints computer name, logged in user, Windows name + version (if available).
This is more for you when reading logs than for grandma.

Reset Windows Update components
Stops services: wuauserv, bits, cryptsvc, msiserver (with friendly OK / warning messages).
Renames:
%windir%\SoftwareDistribution → SoftwareDistribution.bak_timestamp
%windir%\System32\catroot2 → catroot2.bak_timestamp
Restarts the same services.
Everything is wrapped in try/catch with “this is not always a problem” style messages where appropriate.

Repair system files
Runs:
dism.exe /Online /Cleanup-Image /RestoreHealthsfc.exe /scannow
Explains that these can take 10–30 minutes and that it’s normal for the PC to feel slow.
Interprets basic exit codes and prints friendly summaries.

Prepare PSWindowsUpdate
Ensures TLS 1.2 is allowed for older systems.
Tries to install:
NuGet package provider.
PSWindowsUpdate module from PSGallery.

Imports PSWindowsUpdate.
If this fails for corporate lock-down or no internet, etc., it prints:
“This often happens on work computers managed by an IT department or if there is no Internet connection.”
Log path so the user can send you the transcript.
Then exits without attempting updates.

Ask Windows for updates
Runs:
Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Verbose
Prints whether updates were found and processed, and tells the user to scroll up to see details and any “restart required” messages.

Final instructions
Summarises next steps:
Restart if required.
Check Settings -> Windows Update after reboot.
Send you the log file if errors persist.

Logging and output files
All runs create a transcript in a Desktop log folder:
Folder:
Desktop\WindowsUpdateHelper_Logs\
Per-run log file:
WindowsUpdateHelper_yyyyMMdd_HHmmss.txt
No other files are created next to the script; update cache backups live in %windir%.

**Limitations / When not to use**
Corporate / domain-managed machines
Group Policy, WSUS, or security baselines may block module installs or updates. The script will say so in a friendly way but cannot bypass admin policies.
Serious corruption or disk errors
DISM/SFC can only do so much. If both fail repeatedly, you’re probably looking at heavier repair or reinstall.
Third-party AV / endpoint protection
Some endpoint tools can interfere with updates or PowerShell. The script does not try to fight those.
This is not a “hack everything” tool. It’s a repeatable good first step for home PCs.

**Troubleshooting**
Script says it must be run as administrator
User did not right-click and choose “Run as administrator”. Have them try again and confirm the UAC prompt.
Stuck for a long time on DISM or SFC
Normal for slow machines or HDDs. As long as the percentage changes occasionally, let it run. If it’s frozen for an hour, you may want to cancel and check disks.
PSWindowsUpdate cannot be installed / imported
Usually: No internet, or PowerShell Gallery blocked by policy. In that case the script tells the user to send you the log.
Updates still fail after running the script
Ask for: 
The latest WindowsUpdateHelper_*.txt from the Desktop.
A screenshot of the Windows Update error message.
Then you can take over with more targeted tools.

**Intent & License**
Personal helper for my own Windows support life:
“Before I remote in, please run this and send me the log.”
Provided as-is, without warranty. Use at your own risk.
Feel free to fork, trim, or extend it for your own use case.