<#
.SYNOPSIS
    Installs standard staff software on a new Windows 11 (x64) laptop.
    - Writes a transcript to C:\Temp\InstallLogs
    - Drops ".done" files in C:\Temp\InstallMarkers so the script is resumable
#>

# --- Safety ----------------
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Logging ---------------
$logRoot = 'C:\Temp\InstallLogs'
if (-not (Test-Path $logRoot)) { New-Item -ItemType Directory -Path $logRoot | Out-Null }
$logFile = Join-Path $logRoot ("Install-{0:yyyyMMdd-HHmm}.log" -f (Get-Date))
Start-Transcript -Path $logFile -Append

# --- Marker-files root -------------
$markerRoot = 'C:\Temp\InstallMarkers'
if (-not (Test-Path $markerRoot)) { New-Item -ItemType Directory -Path $markerRoot | Out-Null }

# --- Helper: copy + unblock -----------
function Copy-Unblock {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$TargetPath
    )
    Copy-Item $SourcePath $TargetPath -Force
    Unblock-File $TargetPath
}

# --- Helper: resumable installer step ---------
function Install-Step {
    param(
        [string]      $Name,
        [scriptblock] $Action,
        [string]      $MarkerFile
    )

    if (Test-Path $MarkerFile) {
        Write-Host "$Name : already completed, skipping." -ForegroundColor DarkYellow
        return
    }

    Write-Host "---- $Name ----" -ForegroundColor Cyan
    try {
        & $Action
        New-Item -ItemType File -Path $MarkerFile -Force | Out-Null
        Write-Host "$Name : OK`n" -ForegroundColor Green
    }
    catch {
        Write-Warning "$Name : FAILED - $($_.Exception.Message)"
        throw   # stop the script so we can re-run after fixing
    }
}

# --- 1. Office 2019 ------------
Install-Step "Copy Officeinstall folder" {
    robocopy "\\trinidad\GroupPolicyInstallSoftware\OFFICE 2019\Officeinstall" `
             "C:\Officeinstall" /E /NFL /NDL /NJH /NJS /NC
} -MarkerFile "$markerRoot\office-copied.done"

Install-Step "Run Office setup /configure" {
    Push-Location "C:\Officeinstall"
    Start-Process ".\Setup.exe" -ArgumentList "/configure configuration.xml" -Wait
    Pop-Location
} -MarkerFile "$markerRoot\office-installed.done"

# --- 2. Fonts4Teachers -------------
Install-Step "Install Fonts4Teachers" {
    $src = "\\10.16.2.44\Install\Fonts4Teachers Deluxe\fonts4teachers_deluxe.exe"
    $dst = "$env:TEMP\fonts4teachers_deluxe.exe"
    Copy-Unblock $src $dst
    Start-Process $dst -ArgumentList "/S" -Wait
    Remove-Item $dst
} -MarkerFile "$markerRoot\fonts4teachers.done"

# --- 3. SMART Notebook -------------
Install-Step "Install SMART Notebook" {
    $src = "\\10.16.2.44\Install\Smart Notebook\smart24-1web.exe"
    $dst = "$env:TEMP\smart24-1web.exe"
    Copy-Unblock $src $dst

    # Run silently, suppress auto-reboot
    $proc = Start-Process $dst -ArgumentList "/quiet REBOOT=ReallySuppress" -Wait -PassThru
    if ($proc.ExitCode -eq 3010) { $global:RebootNeeded = $true }

    Remove-Item $dst
} -MarkerFile "$markerRoot\smart-notebook.done"

# --- 4. Acrobat Reader (no McAfee) ---------
Install-Step "Install Acrobat Reader" {
    $url    = 'https://download.adobe.com/pub/adobe/reader/win/AcrobatDC/2300420588/AcroRdrDCx642400420588_en_US.exe'
    $target = "$env:TEMP\AcroRdr.exe"
    Invoke-WebRequest $url -OutFile $target
    Start-Process $target -ArgumentList "/sAll /rs /rps /msi EULA_ACCEPT=YES REMOVE_PREVIOUS=YES" -Wait
    Remove-Item $target
} -MarkerFile "$markerRoot\acrobat.done"

# --- 5. VLC Media Player -----------
Install-Step "Install VLC" {
    $url    = 'https://downloads.videolan.org/pub/videolan/vlc/3.0.20/win64/vlc-3.0.20-win64.exe'
    $target = "$env:TEMP\VLC.exe"
    Invoke-WebRequest $url -OutFile $target
    Start-Process $target -ArgumentList "/S" -Wait
    Remove-Item $target
} -MarkerFile "$markerRoot\vlc.done"

# --- 6. Google Chrome Enterprise -----------
Install-Step "Install Chrome" {
    $url    = 'https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D&iid=&lang=en&browser=5&usagestats=0&appname=Google%20Chrome&type=standalone-64/edgedl/chrome/install/GoogleChromeStandaloneEnterprise64.msi'
    $target = "$env:TEMP\Chrome.msi"
    Invoke-WebRequest $url -OutFile $target
    Start-Process msiexec.exe -ArgumentList "/i `"$target`" /qn /norestart" -Wait
    Remove-Item $target
} -MarkerFile "$markerRoot\chrome.done"

# --- Finish up -------------
Stop-Transcript
Write-Host "`nAll steps finished. Log: $logFile" -ForegroundColor Yellow

if ($RebootNeeded) {
    Write-Host "Installations finished - system will restart in 60 seconds..."
    shutdown /r /t 60
}
