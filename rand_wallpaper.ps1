
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
  Include images from subfolders. Default: false.

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
  [string]$ImageFolder,

  [int]$IntervalMinutes = 30,

  [ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')]
  [string]$Style = 'Fill',

  [switch]$Recurse,

  [switch]$Once,

  [switch]$InstallBackground,

  [switch]$UninstallBackground,

  [string]$TaskName = 'RandomWallpaper',

  [switch]$LockScreen
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

function Get-LockScreenTargetPath {
  # Prefer system path for lock screen assets; fallback to ProgramData if not writable
  $folder = 'C:\Windows\Web\Screen'
  try {
    if (-not (Test-Path -LiteralPath $folder)) {
      New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
  } catch {
    $folder = 'C:\ProgramData\RandomWallpaper'
    if (-not (Test-Path -LiteralPath $folder)) {
      New-Item -Path $folder -ItemType Directory -Force | Out-Null
    }
  }
  return (Join-Path $folder 'RandomWallpaper.jpg')
}

function Convert-ToJpeg {
  param(
    [Parameter(Mandatory)][string]$SourcePath,
    [Parameter(Mandatory)][string]$DestPath,
    [int]$Quality = 85
  )
  Add-Type -AssemblyName System.Drawing
  $img = $null
  try {
    $img = [System.Drawing.Image]::FromFile($SourcePath)
    $jpgCodec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $encParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [int64]$Quality)
    if (Test-Path -LiteralPath $DestPath) { Remove-Item -LiteralPath $DestPath -Force -ErrorAction SilentlyContinue }
    $img.Save($DestPath, $jpgCodec, $encParams)
  } finally {
    if ($img) { $img.Dispose() }
  }
}

function Ensure-FileReadableByUsers {
  param([Parameter(Mandatory)][string]$Path)
  try {
    $acl = Get-Acl -LiteralPath $Path
    # Use SIDs to avoid localization issues
    $sidUsers = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-545')            # BUILTIN\Users
    $sidSystem = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')               # LOCAL SYSTEM
    $sidAllApps = New-Object System.Security.Principal.SecurityIdentifier('S-1-15-2-1')            # ALL APPLICATION PACKAGES
    $rights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::Read
    $inherit = [System.Security.AccessControl.InheritanceFlags]::None
    $prop = [System.Security.AccessControl.PropagationFlags]::None
    $allow = [System.Security.AccessControl.AccessControlType]::Allow
    $rules = @(
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidUsers, $rights, $inherit, $prop, $allow),
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem, $rights, $inherit, $prop, $allow),
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidAllApps, $rights, $inherit, $prop, $allow)
    )
    foreach ($rule in $rules) { $acl.AddAccessRule($rule) }
    Set-Acl -LiteralPath $Path -AclObject $acl
  } catch { }
}

function Ensure-FolderReadableByUsers {
  param([Parameter(Mandatory)][string]$FolderPath)
  try {
    $acl = Get-Acl -LiteralPath $FolderPath
    $sidUsers = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-32-545')
    $sidSystem = New-Object System.Security.Principal.SecurityIdentifier('S-1-5-18')
    $sidAllApps = New-Object System.Security.Principal.SecurityIdentifier('S-1-15-2-1')
    $rights = [System.Security.AccessControl.FileSystemRights]::ReadAndExecute -bor [System.Security.AccessControl.FileSystemRights]::ListDirectory
    $inherit = [System.Security.AccessControl.InheritanceFlags]::ContainerInherit -bor [System.Security.AccessControl.InheritanceFlags]::ObjectInherit
    $prop = [System.Security.AccessControl.PropagationFlags]::None
    $allow = [System.Security.AccessControl.AccessControlType]::Allow
    $rules = @(
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidUsers, $rights, $inherit, $prop, $allow),
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidSystem, $rights, $inherit, $prop, $allow),
      New-Object System.Security.AccessControl.FileSystemAccessRule($sidAllApps, $rights, $inherit, $prop, $allow)
    )
    foreach ($rule in $rules) { $acl.AddAccessRule($rule) }
    Set-Acl -LiteralPath $FolderPath -AclObject $acl
  } catch { }
}

function Set-LockScreenImage {
  param([Parameter(Mandatory)][string]$ImagePath)
  $targetPath = Get-LockScreenTargetPath
  $targetFolder = Split-Path -Path $targetPath -Parent

  # Ensure folder readable by system components
  Ensure-FolderReadableByUsers -FolderPath $targetFolder

  # Always re-encode to real JPEG to satisfy policy expectations
  try {
    Convert-ToJpeg -SourcePath $ImagePath -DestPath $targetPath -Quality 85
    Ensure-FileReadableByUsers -Path $targetPath
  } catch {
    throw "Failed to prepare JPEG for lock screen at ${targetPath}: $($_.Exception.Message)"
  }

  if (Test-IsAdmin) {
    try {
      # Set classic policy key for compatibility
      New-Item -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization' -Force | Out-Null
      Set-ItemProperty -Path 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Personalization' -Name 'LockScreenImage' -Type String -Value $targetPath

      # Set PersonalizationCSP keys (device-wide) per request
      New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Force | Out-Null
      Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImagePath' -Type String -Value $targetPath
      Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImageUrl'  -Type String -Value $targetPath
      Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImageStatus' -Type DWord  -Value 1

      # Apply policy (best-effort)
      try { gpupdate /target:computer /force | Out-Null } catch {}
      Write-Host "Lock screen image policy and CSP set to $targetPath"
    } catch {
      Write-Warning "Failed to set policy/CSP lock screen image: $($_.Exception.Message)"
    }
  } else {
    # Best-effort user-level approach (may not work on all editions)
    try {
      Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' -Name 'RotatingLockScreenEnabled' -Type DWord -Value 0 -ErrorAction SilentlyContinue
      Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager' -Name 'RotatingLockScreenOverlayEnabled' -Type DWord -Value 0 -ErrorAction SilentlyContinue
      New-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Force -ErrorAction SilentlyContinue | Out-Null
      Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImagePath' -Type String -Value $targetPath -ErrorAction SilentlyContinue
      Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImageUrl' -Type String -Value $targetPath -ErrorAction SilentlyContinue
      Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name 'LockScreenImageStatus' -Type DWord -Value 1 -ErrorAction SilentlyContinue
      Write-Host "Attempted to set user lock screen image to $targetPath (may require MDM/admin)."
    } catch {
      Write-Warning "Failed to apply user-level lock screen settings: $($_.Exception.Message)"
    }
  }
}

function Set-RandomWallpaperOnce {
  param([string]$Folder, [switch]$IncludeSubfolders, [string]$Style)
  $image = Get-RandomImagePath -Folder $Folder -IncludeSubfolders:$IncludeSubfolders
  Write-Host "Setting wallpaper: $image (Style: $Style)"
  Set-WallpaperStyleRegistry -Style $Style
  Apply-Wallpaper -ImagePath $image
  if ($LockScreen) {
    try { Set-LockScreenImage -ImagePath $image } catch { Write-Warning $_ }
  }
}

function Install-BackgroundTask {
  param(
    [Parameter(Mandatory)][string]$ScriptPath,
    [Parameter(Mandatory)][string]$ResolvedImageFolder,
    [int]$IntervalMinutes = 30,
    [string]$Style = 'Fill',
    [switch]$Recurse,
    [switch]$LockScreen,
    [string]$TaskName = 'RandomWallpaper'
  )
  Write-Host "Installing Scheduled Task '$TaskName' to run in background..."

  # Build arguments to re-run this script without -Once so it loops continuously
  $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',("`"{0}`"" -f $ScriptPath))
  if ($ResolvedImageFolder) { $argList += @('-ImageFolder',("`"{0}`"" -f $ResolvedImageFolder)) }
  if ($IntervalMinutes)   { $argList += @('-IntervalMinutes',("{0}" -f $IntervalMinutes)) }
  if ($Style)             { $argList += @('-Style',("{0}" -f $Style)) }
  if ($Recurse)           { $argList += '-Recurse' }
  if ($LockScreen)        { $argList += '-LockScreen' }

  $action   = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument ($argList -join ' ')
  $trigger1 = New-ScheduledTaskTrigger -AtLogOn

  # Optional repeating trigger as a watchdog to restart if stopped; main script loops internally
  $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $runLevel = if ($LockScreen) { 'Highest' } else { 'Limited' }
  $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel $runLevel
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
  Install-BackgroundTask -ScriptPath $scriptPath -ResolvedImageFolder $resolvedFolderForTask -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse -LockScreen:$LockScreen -TaskName $TaskName
  return
}

# Resolve folder and do an immediate change on start (per requirement)
$resolvedFolder = Resolve-DefaultImageFolder -InputFolder $ImageFolder
Set-RandomWallpaperOnce -Folder $resolvedFolder -IncludeSubfolders:$Recurse -Style $Style

if ($Once) {
  return
}

# Loop for continuous changes
if ($IntervalMinutes -lt 1) { $IntervalMinutes = 1 }
while ($true) {
  try {
    Start-Sleep -Seconds ($IntervalMinutes * 60)
    Set-RandomWallpaperOnce -Folder $resolvedFolder -IncludeSubfolders:$Recurse -Style $Style
  } catch {
    Write-Warning ("Failed to set wallpaper: " + $_.Exception.Message)
  }
}
