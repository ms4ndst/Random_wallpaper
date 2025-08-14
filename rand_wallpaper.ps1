
<#
.SYNOPSIS
  Set a random desktop wallpaper from a folder. Immediately sets on start, then optionally loops on an interval.

.DESCRIPTION
  - Selects a random image from a specified folder (optionally including subfolders)
  - Sets registry values for wallpaper style (Fill, Fit, Stretch, Center, Tile, Span)
  - Applies the wallpaper using SystemParametersInfo (works on Windows 10/11)
  - Immediately sets a new wallpaper when the script begins
  - Optionally runs continuously and changes the wallpaper on a timer

.PARAMETER ImageFolder
  Folder containing images. Defaults to "Pictures\Wallpapers" under the current user profile if it exists,
  otherwise falls back to the user's Pictures folder.

.PARAMETER IntervalMinutes
  Minutes between wallpaper changes when running continuously. Default 30.

.PARAMETER Style
  Wallpaper style: Fill, Fit, Stretch, Center, Tile, Span. Default Fill.

.PARAMETER Recurse
  Include images from subfolders. Default: true.

.PARAMETER Once
  If specified, set a random wallpaper once and exit (no loop).

.EXAMPLE
  # Set one random wallpaper from the default folder and exit
  ./rand_wallpaper.ps1 -Once

.EXAMPLE
  # Run continuously every 15 minutes, searching subfolders, using Fit style from a custom folder
  ./rand_wallpaper.ps1 -ImageFolder "C:\Wallpapers" -IntervalMinutes 15 -Recurse -Style Fit

.NOTES
  Requires pwsh/PowerShell 5+ on Windows 10/11. Supported image types: jpg, jpeg, png, bmp.
#>
param(
  [Parameter(Position = 0)]
  [string]$ImageFolder = (Join-Path $env:USERPROFILE 'wallpapers'),

  [int]$IntervalMinutes = 30,

  [ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')]
  [string]$Style = 'Stretch',

  [switch]$Recurse,

  [switch]$Once,

  [switch]$InstallBackground,

  [switch]$UninstallBackground,

  [string]$TaskName = 'RandomWallpaper'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-DefaultImageFolder {
  param([string]$InputFolder)
  if ($InputFolder -and (Test-Path -LiteralPath $InputFolder)) { return (Resolve-Path -LiteralPath $InputFolder).Path }

  # Preferred default: %USERPROFILE%\wallpapers
  $userWall = Join-Path $env:USERPROFILE 'wallpapers'
  if (Test-Path -LiteralPath $userWall) { return (Resolve-Path -LiteralPath $userWall).Path }

  # Fallbacks: Pictures\Wallpapers, then Pictures
  $pictures = Join-Path $env:USERPROFILE 'Pictures'
  $defaultWall = Join-Path $pictures 'Wallpapers'
  if (Test-Path -LiteralPath $defaultWall) { return (Resolve-Path -LiteralPath $defaultWall).Path }
  if (Test-Path -LiteralPath $pictures) { return (Resolve-Path -LiteralPath $pictures).Path }

  throw "No valid image folder found. Provide -ImageFolder with a valid path."
}

function Get-RandomImagePath {
  param(
    [Parameter(Mandatory)] [string]$Folder,
    [switch]$IncludeSubfolders
  )
  $patterns = @('*.jpg','*.jpeg','*.png','*.bmp')
  $files = foreach ($pattern in $patterns) {
    Get-ChildItem -LiteralPath $Folder -Filter $pattern -File -ErrorAction SilentlyContinue -Recurse:$IncludeSubfolders
  }
  if (-not $files -or $files.Count -eq 0) {
    throw "No images found in '$Folder'$(if($IncludeSubfolders){' (including subfolders)'}). Supported: jpg, jpeg, png, bmp."
  }
  ($files | Get-Random).FullName
}

function Set-WallpaperStyleRegistry {
  param([ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')] [string]$Style)
  $desktopKey = 'HKCU:\Control Panel\Desktop'
  switch ($Style) {
    'Fill'    { $wallpaperStyle = '10'; $tile = '0' }
    'Fit'     { $wallpaperStyle = '6';  $tile = '0' }
    'Stretch' { $wallpaperStyle = '2';  $tile = '0' }
    'Center'  { $wallpaperStyle = '0';  $tile = '0' }
    'Tile'    { $wallpaperStyle = '0';  $tile = '1' }
    'Span'    { $wallpaperStyle = '22'; $tile = '0' } # Multi-monitor span
  }
  Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value $wallpaperStyle -Type String | Out-Null
  Set-ItemProperty -Path $desktopKey -Name TileWallpaper  -Value $tile           -Type String | Out-Null
}

# Add native call to SystemParametersInfo for applying the wallpaper
Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
  [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
  public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
"@

function Apply-Wallpaper {
  param([Parameter(Mandatory)][string]$ImagePath)
  # SPI_SETDESKWALLPAPER = 20, SPIF_UPDATEINIFILE = 0x01, SPIF_SENDWININICHANGE = 0x02
  $SPI_SETDESKWALLPAPER = 20
  $flags = 0x01 -bor 0x02
  $ok = [Win32.NativeMethods]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $ImagePath, $flags)
  if (-not $ok) {
    $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw "SystemParametersInfo failed with error code $err for path: $ImagePath"
  }
}

function Test-IsAdmin {
  try {
    $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wp = New-Object Security.Principal.WindowsPrincipal($wi)
    return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}






function Set-RandomWallpaperOnce {
  param([string]$Folder, [switch]$IncludeSubfolders, [string]$Style)
  $image = Get-RandomImagePath -Folder $Folder -IncludeSubfolders:$IncludeSubfolders
  Write-Host "Setting wallpaper: $image (Style: $Style)"
  Set-WallpaperStyleRegistry -Style $Style
  Apply-Wallpaper -ImagePath $image
}

function Install-BackgroundTask {
  param(
    [Parameter(Mandatory)][string]$ScriptPath,
    [Parameter(Mandatory)][string]$ResolvedImageFolder,
    [int]$IntervalMinutes = 30,
    [string]$Style = 'Fill',
    [switch]$Recurse,
    [string]$TaskName = 'RandomWallpaper'
  )
  Write-Host "Installing Scheduled Task '$TaskName' to run in background..."

  # Build arguments to re-run this script without -Once so it loops continuously
  $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',("`"{0}`"" -f $ScriptPath))
  if ($ResolvedImageFolder) { $argList += @('-ImageFolder',("`"{0}`"" -f $ResolvedImageFolder)) }
  if ($IntervalMinutes)   { $argList += @('-IntervalMinutes',("{0}" -f $IntervalMinutes)) }
  if ($Style)             { $argList += @('-Style',("{0}" -f $Style)) }
  if ($Recurse)           { $argList += '-Recurse' }

  $action   = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument ($argList -join ' ')
  $trigger1 = New-ScheduledTaskTrigger -AtLogOn

  # Optional repeating trigger as a watchdog to restart if stopped; main script loops internally
  $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
  $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger1 -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable)

  try {
    Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    Write-Host "Scheduled Task '$TaskName' installed. Starting it now..."
    Start-ScheduledTask -TaskName $TaskName
    Write-Host "Scheduled Task '$TaskName' started."
  } catch {
    throw "Failed to register/start Scheduled Task '$TaskName': $($_.Exception.Message)"
  }
}

function Uninstall-BackgroundTask {
  param([string]$TaskName = 'RandomWallpaper')
  try {
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
      Write-Host "Stopping and removing Scheduled Task '$TaskName'..."
      try { Stop-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue } catch {}
      Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
      Write-Host "Scheduled Task '$TaskName' removed."
    } else {
      Write-Host "Scheduled Task '$TaskName' not found."
    }
  } catch {
    throw "Failed to remove Scheduled Task '$TaskName': $($_.Exception.Message)"
  }
}

# If user requested background install/uninstall, handle that first and exit
$scriptPath = $MyInvocation.MyCommand.Path
if ($UninstallBackground) {
  Uninstall-BackgroundTask -TaskName $TaskName
  return
}
if ($InstallBackground) {
  $resolvedFolderForTask = Resolve-DefaultImageFolder -InputFolder $ImageFolder
  Install-BackgroundTask -ScriptPath $scriptPath -ResolvedImageFolder $resolvedFolderForTask -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse -TaskName $TaskName
  return
}

# Resolve folder and do an immediate change on start (per requirement)
$resolvedFolder = Resolve-DefaultImageFolder -InputFolder $ImageFolder
# Default recursion to true unless explicitly provided
$useRecurse = if ($PSBoundParameters.ContainsKey('Recurse')) { [bool]$Recurse } else { $true }
Set-RandomWallpaperOnce -Folder $resolvedFolder -IncludeSubfolders:$useRecurse -Style $Style

if ($Once) {
  return
}

# Loop for continuous changes
if ($IntervalMinutes -lt 1) { $IntervalMinutes = 1 }
while ($true) {
  try {
    Start-Sleep -Seconds ($IntervalMinutes * 60)
    Set-RandomWallpaperOnce -Folder $resolvedFolder -IncludeSubfolders:$useRecurse -Style $Style
  } catch {
    Write-Warning ("Failed to set wallpaper: " + $_.Exception.Message)
  }
}
