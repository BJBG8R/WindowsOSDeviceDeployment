<#
.SYNOPSIS
    NUO Device Deployment Script

.DESCRIPTION
    Configures a Windows device by clearing taskbar pins, collecting user
    information, configuring the preferred browser, adding NUO bookmarks,
    installing Microsoft 365 Copilot, installing NaviGator Chat as a web app,
    pinning applications, configuring Remote Desktop or laptop RDP access,
    and optionally configuring biometric sign-in.

.NOTES
    Run this script as Administrator.
#>

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# ERROR HANDLING
# ---------------------------------------------------------------------------

trap {
    Write-Host ""
    Write-Host "SCRIPT FAILED" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to close this window" | Out-Null
    exit 1
}

# ---------------------------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------------------------

function Pause-Script {
    param(
        [string]$Message = "Press Enter to continue..."
    )

    Write-Host ""
    Read-Host $Message | Out-Null
}

function Get-ApplicationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Paths
    )

    foreach ($path in $Paths) {
        if ($path -and (Test-Path $path)) {
            return $path
        }
    }

    return $null
}

function Get-PreferredBrowserPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BrowserChoice
    )

    switch ($BrowserChoice) {
        "1" {
            $command = Get-Command "chrome.exe" -ErrorAction SilentlyContinue

            if ($command) {
                return $command.Source
            }

            return Get-ApplicationPath -Paths @(
                "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
                "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
                "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe"
            )
        }

        "2" {
            $command = Get-Command "firefox.exe" -ErrorAction SilentlyContinue

            if ($command) {
                return $command.Source
            }

            return Get-ApplicationPath -Paths @(
                "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
                "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
            )
        }

        "3" {
            $command = Get-Command "msedge.exe" -ErrorAction SilentlyContinue

            if ($command) {
                return $command.Source
            }

            return Get-ApplicationPath -Paths @(
                "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
                "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
                "$env:LOCALAPPDATA\Microsoft\Edge\Application\msedge.exe"
            )
        }
    }

    return $null
}

function Pin-AppToTaskbar {
    param(
        [Parameter(Mandatory = $true)]
        [string]$AppName
    )

    try {
        $shell = New-Object -ComObject Shell.Application

        $programs = $shell.NameSpace(
            "shell:::{4234d49b-0245-4df3-b780-3893943456e1}"
        )

        if ($null -eq $programs) {
            Write-Warning "Could not access the Start menu applications folder."
            return
        }

        foreach ($item in @($programs.Items())) {
            $itemsToCheck = @($item)

            try {
                $folder = $item.GetFolder

                if ($null -ne $folder) {
                    $itemsToCheck += @($folder.Items())
                }
            }
            catch {
                # Some Start menu entries do not expose folders.
            }

            foreach ($candidate in $itemsToCheck) {
                if ($candidate.Name -like "*$AppName*") {
                    foreach ($verb in @($candidate.Verbs())) {
                        if ($verb.Name -match "Pin to tas&kbar|Pin to Taskbar") {
                            $verb.DoIt()

                            Write-Host `
                                "  [+] Pinned '$AppName' to the taskbar." `
                                -ForegroundColor Green

                            return
                        }
                    }
                }
            }
        }

        Write-Warning `
            "Could not find '$AppName' in the Start menu. Pin it manually if necessary."
    }
    catch {
        Write-Warning `
            "Could not pin '$AppName': $($_.Exception.Message)"
    }
}

function Set-DefaultBrowser {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("1", "2", "3")]
        [string]$BrowserChoice
    )

    switch ($BrowserChoice) {
        "1" {
            $browserName = "Google Chrome"
            $registeredApp = "Google.Chrome"
        }

        "2" {
            $browserName = "Mozilla Firefox"
            $registeredApp = "Mozilla.Firefox"
        }

        "3" {
            $browserName = "Microsoft Edge"
            $registeredApp = "Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge"
        }
    }

    $settingsUri = "ms-settings:defaultapps?registeredAppUser=$registeredApp"

    Write-Host ""
    Write-Host "Opening Default Apps for $browserName..." `
        -ForegroundColor Magenta

    Start-Process -FilePath $settingsUri

    Write-Host ""
    Write-Host "ACTION REQUIRED: Click 'Set default' for $browserName." `
        -ForegroundColor Yellow

    Pause-Script `
        "Press Enter after clicking 'Set default' for $browserName..."
}

function Install-Microsoft365Copilot {
    $downloadUrl =
        "https://go.microsoft.com/fwlink/?linkid=2325486"

    $installDirectory =
        Join-Path $env:TEMP "M365CopilotInstall"

    $installerPath =
        Join-Path $installDirectory "M365CopilotDesktopInstaller.exe"

    Write-Host `
        "  Preparing the Microsoft 365 Copilot installer..." `
        -ForegroundColor Yellow

    New-Item `
        -Path $installDirectory `
        -ItemType Directory `
        -Force `
        | Out-Null

    if (Test-Path $installerPath) {
        Remove-Item `
            -Path $installerPath `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Write-Host `
        "  Downloading Microsoft 365 Copilot..." `
        -ForegroundColor Yellow

    Invoke-WebRequest `
        -Uri $downloadUrl `
        -OutFile $installerPath `
        -UseBasicParsing

    if (-not (Test-Path $installerPath)) {
        throw "The Microsoft 365 Copilot installer was not downloaded."
    }

    Write-Host `
        "  Running the Microsoft 365 Copilot installer..." `
        -ForegroundColor Yellow

    $process = Start-Process `
        -FilePath $installerPath `
        -ArgumentList @(
            "--quiet",
            "--start",
            "-p"
        ) `
        -Wait `
        -PassThru

    if ($process.ExitCode -ne 0) {
        throw `
            "Microsoft 365 Copilot installation failed with exit code $($process.ExitCode)."
    }

    Write-Host `
        "  [+] Microsoft 365 Copilot installation completed." `
        -ForegroundColor Green
}

function Install-NaviGatorChatWebApp {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BrowserChoice,

        [Parameter(Mandatory = $true)]
        [string]$BrowserName
    )

    $naviGatorUrl = "https://chat.ai.it.ufl.edu/"
    $browserPath = Get-PreferredBrowserPath -BrowserChoice $BrowserChoice

    if (-not $browserPath) {
        throw `
            "The user's preferred browser could not be found on this computer."
    }

    Write-Host `
        "  Installing NaviGator Chat as a web app in $BrowserName..." `
        -ForegroundColor Yellow

    switch ($BrowserChoice) {
        "1" {
            # Chrome supports opening a website in app mode.
            # Depending on Chrome policy, the user may need to confirm
            # installation from the browser menu.
            Start-Process `
                -FilePath $browserPath `
                -ArgumentList "--app=$naviGatorUrl"

            Write-Host `
                "  [+] NaviGator Chat opened in Chrome app mode." `
                -ForegroundColor Green
        }

        "2" {
            # Firefox does not support a dependable silent command-line
            # web-app installation method.
            Start-Process `
                -FilePath $browserPath `
                -ArgumentList $naviGatorUrl

            Write-Warning `
                "Firefox does not support reliable command-line web-app installation."

            Write-Host `
                "  Install the site as an app manually from the Firefox menu if needed." `
                -ForegroundColor Magenta
        }

        "3" {
            # Edge supports direct installation of a website as an app.
            Start-Process `
                -FilePath $browserPath `
                -ArgumentList "--install-app=$naviGatorUrl" `
                -Wait

            Write-Host `
                "  [+] NaviGator Chat was installed as an Edge web app." `
                -ForegroundColor Green
        }
    }

    Start-Sleep -Seconds 8

    Pin-AppToTaskbar -AppName "NaviGator Chat"
}

function Add-RemoteDesktopUser {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Username
    )

    $domainUser = "UFAD\$Username"

    try {
        Add-LocalGroupMember `
            -Group "Remote Desktop Users" `
            -Member $domainUser `
            -ErrorAction Stop

        Write-Host `
            "  [+] Added $domainUser to Remote Desktop Users." `
            -ForegroundColor Green
    }
    catch {
        Write-Warning `
            "Could not add $domainUser to Remote Desktop Users: $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# STEP 1 – CLEAR THE TASKBAR
# ---------------------------------------------------------------------------

Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   NUO Device Deployment Script" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 1: Clearing the Taskbar" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

$taskbarDirectory =
    Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Shell"

$taskbarLayoutPath =
    Join-Path $taskbarDirectory "LayoutModification.xml"

New-Item `
    -Path $taskbarDirectory `
    -ItemType Directory `
    -Force `
    | Out-Null

$emptyLayoutLines = @(
    '<?xml version="1.0" encoding="utf-8"?>',
    '<LayoutModificationTemplate',
    '    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"',
    '    xmlns:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"',
    '    xmlns:start="http://schemas.microsoft.com/Start/2014/StartLayout"',
    '    xmlns:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"',
    '    Version="1">',
    '  <CustomTaskbarLayoutCollection PinListPlacement="Replace">',
    '    <defaultlayout:TaskbarLayout>',
    '      <taskbar:TaskbarPinList>',
    '      </taskbar:TaskbarPinList>',
    '    </defaultlayout:TaskbarLayout>',
    '  </CustomTaskbarLayoutCollection>',
    '</LayoutModificationTemplate>'
)

$emptyLayoutLines |
    Out-File `
        -FilePath $taskbarLayoutPath `
        -Encoding UTF8 `
        -Force

Write-Host `
    "  [+] Taskbar layout cleared." `
    -ForegroundColor Green

Stop-Process `
    -Name "explorer" `
    -Force `
    -ErrorAction SilentlyContinue

Start-Sleep -Seconds 3
Start-Process "explorer.exe"
Start-Sleep -Seconds 3

Write-Host `
    "  [+] Explorer restarted." `
    -ForegroundColor Green

# ---------------------------------------------------------------------------
# STEP 2 – USER INFORMATION
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 2: User Information" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

# The user's Gatorlink username is the current Windows username.
$userGatorlink = [Environment]::UserName

$techGatorlink =
    Read-Host "Enter the TECH's Gatorlink username"

$techAdmAccount =
    "adm-$techGatorlink"

Write-Host ""
Write-Host "  User    : $userGatorlink" -ForegroundColor Yellow
Write-Host "  Tech    : $techGatorlink" -ForegroundColor Yellow
Write-Host "  Tech ADM: $techAdmAccount" -ForegroundColor Yellow

# ---------------------------------------------------------------------------
# STEP 3 – PREFERRED BROWSER
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 3: Browser Setup" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

do {
    $browserChoice =
        Read-Host `
            "What is the user's preferred browser? (1=Google Chrome, 2=Firefox, 3=Edge)"
}
while ($browserChoice -notin @("1", "2", "3"))

$browserName = switch ($browserChoice) {
    "1" { "Google Chrome" }
    "2" { "Mozilla Firefox" }
    "3" { "Microsoft Edge" }
}

Set-DefaultBrowser -BrowserChoice $browserChoice

Write-Host `
    "  Pinning $browserName to the taskbar..." `
    -ForegroundColor Yellow

Pin-AppToTaskbar -AppName $browserName

if ($browserChoice -in @("1", "2")) {
    Write-Host ""
    Write-Host `
        "Ask the user to sign in to $browserName synchronization." `
        -ForegroundColor Magenta

    Pause-Script `
        "Press Enter after browser synchronization is configured..."
}

Write-Host ""
Write-Host `
    "In $browserName, open a Zoom meeting link and allow ufl.zoom.us to open Zoom links." `
    -ForegroundColor Magenta

Pause-Script `
    "Press Enter after the Zoom link preference is configured..."

# ---------------------------------------------------------------------------
# STEP 4 – NUO BOOKMARKS
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 4: NUO Bookmarks" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

$pinBookmarks =
    Read-Host `
        "Would you like to add NUO bookmarks to the browser? (Y/N)"

if ($pinBookmarks -match "^[Yy]") {
    $nuoBookmarks = @(
        @{
            Name = "WCB IT Help"
            URL  = "https://wcb.it.ufl.edu"
        }
        @{
            Name = "GatorMail"
            URL  = "https://mail.ufl.edu"
        }
        @{
            Name = "Protect My Email"
            URL  = "https://protectmyemail.ufl.edu"
        }
        @{
            Name = "UFIT Alerts"
            URL  = "https://alerts.it.ufl.edu"
        }
        @{
            Name = "UFIT AI Services"
            URL  = "https://ai.it.ufl.edu"
        }
        @{
            Name = "Fast Path Solutions"
            URL  = "https://fastpath.ufl.edu"
        }
        @{
            Name = "UF Data Guide"
            URL  = "https://dataguide.research.ufl.edu"
        }
        @{
            Name = "Mobile GatorOne"
            URL  = "https://m.gatorone.ufl.edu"
        }
    )

    if ($browserChoice -eq "1") {
        $chromeBookmarksFile =
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Bookmarks"

        if (Test-Path $chromeBookmarksFile) {
            $bookmarksJson =
                Get-Content $chromeBookmarksFile -Raw |
                ConvertFrom-Json

            $existingFolder =
                @($bookmarksJson.roots.bookmark_bar.children) |
                Where-Object {
                    $_.name -eq "NUO" -and $_.type -eq "folder"
                } |
                Select-Object -First 1

            if ($existingFolder) {
                $nuoFolder = $existingFolder
            }
            else {
                $nuoFolder = [PSCustomObject]@{
                    children      = @()
                    date_added    = [string](
                        [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
                    )
                    date_modified = [string](
                        [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
                    )
                    guid          = [guid]::NewGuid().ToString()
                    id            = [string](Get-Random -Minimum 100 -Maximum 999)
                    name          = "NUO"
                    type          = "folder"
                }

                $bookmarksJson.roots.bookmark_bar.children += $nuoFolder
            }

            foreach ($bookmark in $nuoBookmarks) {
                $alreadyExists =
                    @($nuoFolder.children) |
                    Where-Object {
                        $_.url -eq $bookmark.URL
                    }

                if (-not $alreadyExists) {
                    $newBookmark = [PSCustomObject]@{
                        date_added = [string](
                            [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
                        )
                        guid       = [guid]::NewGuid().ToString()
                        id         = [string](Get-Random -Minimum 1000 -Maximum 9999)
                        name       = $bookmark.Name
                        type       = "url"
                        url        = $bookmark.URL
                    }

                    $nuoFolder.children += $newBookmark
                }
            }

            $bookmarksJson |
                ConvertTo-Json -Depth 30 |
                Set-Content `
                    -Path $chromeBookmarksFile `
                    -Encoding UTF8

            Write-Host `
                "  [+] NUO bookmarks added to Chrome." `
                -ForegroundColor Green
        }
        else {
            Write-Warning `
                "Chrome's Bookmarks file was not found."
        }

        $chromePreferencesFile =
            "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Preferences"

        if (Test-Path $chromePreferencesFile) {
            $preferences =
                Get-Content $chromePreferencesFile -Raw |
                ConvertFrom-Json

            if ($null -eq $preferences.bookmark_bar) {
                $preferences |
                    Add-Member `
                        -MemberType NoteProperty `
                        -Name "bookmark_bar" `
                        -Value ([PSCustomObject]@{
                            show_on_all_tabs = $true
                        }) `
                        -Force
            }
            else {
                $preferences.bookmark_bar.show_on_all_tabs = $true
            }

            $preferences |
                ConvertTo-Json -Depth 30 |
                Set-Content `
                    -Path $chromePreferencesFile `
                    -Encoding UTF8

            Write-Host `
                "  [+] Chrome bookmarks toolbar enabled." `
                -ForegroundColor Green
        }
    }
    elseif ($browserChoice -eq "2") {
        Write-Host `
            "Firefox bookmark automation is not enabled." `
            -ForegroundColor Yellow

        foreach ($bookmark in $nuoBookmarks) {
            Write-Host `
                "  $($bookmark.Name): $($bookmark.URL)" `
                -ForegroundColor White
        }

        Pause-Script `
            "Press Enter after adding the Firefox bookmarks manually..."
    }
    elseif ($browserChoice -eq "3") {
        Write-Host `
            "Edge bookmark automation is not enabled." `
            -ForegroundColor Yellow

        foreach ($bookmark in $nuoBookmarks) {
            Write-Host `
                "  $($bookmark.Name): $($bookmark.URL)" `
                -ForegroundColor White
        }

        Pause-Script `
            "Press Enter after adding the Edge bookmarks manually..."
    }
}
else {
    Write-Host `
        "  Skipping NUO bookmarks." `
        -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# STEP 5 – MICROSOFT 365 COPILOT
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 5: Microsoft 365 Copilot" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

Install-Microsoft365Copilot

Start-Sleep -Seconds 5

Pin-AppToTaskbar `
    -AppName "Microsoft 365 Copilot"

# ---------------------------------------------------------------------------
# STEP 6 – NAVIGATOR CHAT
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 6: NaviGator Chat" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

Install-NaviGatorChatWebApp `
    -BrowserChoice $browserChoice `
    -BrowserName $browserName

# ---------------------------------------------------------------------------
# STEP 7 – OUTLOOK PREFERENCE
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 7: Outlook Version" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

do {
    $outlookChoice =
        Read-Host `
            "Does the user prefer Outlook (New) or Outlook (Classic)? (1=New, 2=Classic)"
}
while ($outlookChoice -notin @("1", "2"))

$outlookAppName =
    if ($outlookChoice -eq "1") {
        "Outlook (New)"
    }
    else {
        "Outlook"
    }

Pin-AppToTaskbar -AppName $outlookAppName

# ---------------------------------------------------------------------------
# STEP 8 – PIN ONENOTE AND TEAMS
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 8: Pinning OneNote and Teams" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

foreach ($application in @("OneNote", "Microsoft Teams")) {
    Pin-AppToTaskbar -AppName $application
}

# ---------------------------------------------------------------------------
# STEP 9 – WEBEX
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 9: Webex" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

$usesWebex =
    Read-Host "Does the user use Webex? (Y/N)"

if ($usesWebex -match "^[Yy]") {
    Write-Host `
        "Install Webex from Software Center." `
        -ForegroundColor Magenta

    Pause-Script `
        "Press Enter after Webex has been installed..."

    Pin-AppToTaskbar -AppName "Webex"
}
else {
    Write-Host `
        "  Skipping Webex." `
        -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# STEP 10 – PIN ZOOM
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 10: Pinning Zoom Workplace" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

Pin-AppToTaskbar -AppName "Zoom Workplace"

# ---------------------------------------------------------------------------
# STEP 11 – DEVICE TYPE CONFIGURATION
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 11: Device Type Configuration" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

do {
    $deviceType =
        Read-Host "Is this a Desktop or Laptop? (1=Desktop, 2=Laptop)"
}
while ($deviceType -notin @("1", "2"))

if ($deviceType -eq "1") {
    Write-Host ""
    Write-Host `
        "  Enabling Remote Desktop..." `
        -ForegroundColor Yellow

    Set-ItemProperty `
        -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" `
        -Name "fDenyTSConnections" `
        -Value 0

    Enable-NetFirewallRule `
        -DisplayGroup "Remote Desktop"

    Write-Host `
        "  [+] Remote Desktop enabled." `
        -ForegroundColor Green

    Write-Host `
        "  Adding Remote Desktop users..." `
        -ForegroundColor Yellow

    Add-RemoteDesktopUser -Username $userGatorlink
    Add-RemoteDesktopUser -Username $techGatorlink
    Add-RemoteDesktopUser -Username $techAdmAccount

    $computerName =
        $env:COMPUTERNAME

    $rdpFilePath =
        Read-Host `
            "Enter the full path for the RDP file"

    $rdpFileLines = @(
        "screen mode id:i:2",
        "use multimon:i:0",
        "desktopwidth:i:1920",
        "desktopheight:i:1080",
        "session bpp:i:32",
        "compression:i:1",
        "keyboardhook:i:2",
        "audiocapturemode:i:0",
        "videoplaybackmode:i:1",
        "connection type:i:7",
        "networkautodetect:i:1",
        "bandwidthautodetect:i:1",
        "displayconnectionbar:i:1",
        "enableworkspacereconnect:i:0",
        "disable wallpaper:i:0",
        "allow font smoothing:i:0",
        "disable full window drag:i:1",
        "disable menu anims:i:1",
        "bitmapcachepersistenable:i:1",
        "full address:s:$computerName.ad.ufl.edu",
        "username:s:UFAD\$userGatorlink",
        "audiomode:i:0",
        "redirectprinters:i:1",
        "redirectclipboard:i:1",
        "autoreconnection enabled:i:1",
        "authentication level:i:2",
        "prompt for credentials:i:0",
        "negotiate security layer:i:1"
    )

    $rdpFileLines |
        Out-File `
            -FilePath $rdpFilePath `
            -Encoding ASCII

    Write-Host `
        "  [+] RDP file saved to $rdpFilePath" `
        -ForegroundColor Green
}
else {
    Write-Host ""
    Write-Host `
        "  Pinning Cisco Secure Client..." `
        -ForegroundColor Yellow

    Pin-AppToTaskbar -AppName "Cisco Secure Client"

    $desktopComputerName =
        Read-Host `
            "Enter the user's Warrington-managed desktop computer name"

    $temporaryRdpPath =
        Join-Path $env:TEMP "NUO_RDP_Setup.rdp"

    $laptopRdpLines = @(
        "full address:s:$desktopComputerName.ad.ufl.edu",
        "username:s:UFAD\$userGatorlink",
        "prompt for credentials:i:1"
    )

    $laptopRdpLines |
        Out-File `
            -FilePath $temporaryRdpPath `
            -Encoding ASCII

    Start-Process `
        -FilePath "mstsc.exe" `
        -ArgumentList ('"{0}"' -f $temporaryRdpPath)

    Write-Host `
        "  [+] Remote Desktop Connection launched." `
        -ForegroundColor Green

    Pause-Script `
        "Press Enter after the Remote Desktop window opens..."
}

# ---------------------------------------------------------------------------
# STEP 12 – BIOMETRIC SIGN-IN
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "--------------------------------------------" -ForegroundColor Cyan
Write-Host " Step 12: Biometric Sign-In" -ForegroundColor Cyan
Write-Host "--------------------------------------------" -ForegroundColor Cyan

$biometric =
    Read-Host `
        "Would the user like facial or fingerprint recognition? (Y/N)"

if ($biometric -match "^[Yy]") {
    Start-Process "ms-settings:signinoptions"

    Write-Host `
        "  [+] Sign-in options opened." `
        -ForegroundColor Green
}
else {
    Write-Host `
        "  Skipping biometric setup." `
        -ForegroundColor Yellow
}

# ---------------------------------------------------------------------------
# COMPLETE
# ---------------------------------------------------------------------------

Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "   Device Deployment Complete!" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""

Read-Host `
    "Deployment finished. Press Enter to close this window" `
    | Out-Null