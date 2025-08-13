<#
.SYNOPSIS
  Creates a Startup entry that launches rand_wallpaper.ps1 at user logon with specified options.

.DESCRIPTION
  - Tries to create a .lnk shortcut in the user's Startup folder using WScript.Shell
  - Falls back to creating a .cmd launcher if COM is unavailable
  - Points to pwsh.exe if available, otherwise powershell.exe
  - Passes through Interval/Style/Recurse/LockScreen options to the wallpaper script

.PARAMETER ScriptPath
  Path to rand_wallpaper.ps1. Defaults to rand_wallpaper.ps1 in the same directory as this script.

.PARAMETER IntervalMinutes
  Interval in minutes between changes. Default 15.

.PARAMETER Style
  Fill, Fit, Stretch, Center, Tile, Span. Default Fill.

.PARAMETER Recurse
  Include subfolders when picking images.

.PARAMETER LockScreen
  Also set lock screen image (best-effort unless running as admin).

.PARAMETER TaskName
  Used to name files created; default "RandomWallpaper"

.EXAMPLE
  # Create Startup entry using defaults
  ./Create-RandomWallpaperStartup.ps1 -IntervalMinutes 15 -Style Fill -Recurse -LockScreen
#>
param(
  [string]$ScriptPath = (Join-Path $PSScriptRoot 'rand_wallpaper.ps1'),
  [int]$IntervalMinutes = 15,
  [ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')]
  [string]$Style = 'Fill',
  [switch]$Recurse,
  [switch]$LockScreen = $True,
  [string]$TaskName = 'RandomWallpaper'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-PreferredPsExeForBackground {
  # Prefer Windows PowerShell (powershell.exe) for reliable -WindowStyle Hidden support
  try { return (Get-Command powershell -ErrorAction Stop).Source } catch {}
  try { return (Get-Command pwsh -ErrorAction Stop).Source } catch {}
  throw 'Neither pwsh.exe nor powershell.exe was found in PATH.'
}

function Get-StartupFolder {
  $startup = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Startup'
  if (-not (Test-Path -LiteralPath $startup)) {
    throw "Startup folder not found: $startup"
  }
  return $startup
}

function New-StartupShortcut {
  param(
    [Parameter(Mandatory)][string]$ShortcutPath,
    [Parameter(Mandatory)][string]$TargetPath,
    [Parameter(Mandatory)][string]$Arguments,
    [Parameter(Mandatory)][string]$WorkingDirectory
  )
  try {
    $shell = New-Object -ComObject WScript.Shell
    $sc = $shell.CreateShortcut($ShortcutPath)
    $sc.TargetPath = $TargetPath
    $sc.Arguments = $Arguments
    $sc.WorkingDirectory = $WorkingDirectory
    $sc.IconLocation = $TargetPath
    # 7 = Minimized; true hidden is not supported via .lnk, so we use Hidden flag in arguments too
    $sc.WindowStyle = 7
    $sc.Save()
    if (-not (Test-Path -LiteralPath $ShortcutPath)) {
      throw "Shortcut did not get created at $ShortcutPath"
    }
    return $true
  } catch {
    Write-Warning ("Failed to create .lnk shortcut: " + $_.Exception.Message)
    return $false
  }
}

function New-StartupCmdLauncher {
  param(
    [Parameter(Mandatory)][string]$CmdPath,
    [Parameter(Mandatory)][string]$ScriptPath,
    [Parameter(Mandatory)][string]$Args
  )
  $lines = @(
    '@echo off',
    ('start "" powershell -NoProfile -ExecutionPolicy Bypass -File "{0}" {1}' -f $ScriptPath, $Args)
  )
  Set-Content -LiteralPath $CmdPath -Value $lines -Encoding ASCII
  if (-not (Test-Path -LiteralPath $CmdPath)) {
    throw "Failed to create startup launcher at $CmdPath"
  }
}

# Resolve absolute script path
$scriptFull = (Resolve-Path -LiteralPath $ScriptPath).Path
$startupFolder = Get-StartupFolder
$psExe = Get-PreferredPsExeForBackground

# Build arguments to pass to PowerShell (prefer powershell.exe) and hide/minimize window
$argList = @('-NoProfile','-WindowStyle','Hidden','-ExecutionPolicy','Bypass','-File', ('"{0}"' -f $scriptFull), '-IntervalMinutes', ('{0}' -f $IntervalMinutes), '-Style', $Style)
if ($Recurse) { $argList += '-Recurse' }
if ($LockScreen) { $argList += '-LockScreen' }
$joinedArgs = ($argList -join ' ')

# Try .lnk first
$shortcutPath = Join-Path $startupFolder ($TaskName + '.lnk')
$created = New-StartupShortcut -ShortcutPath $shortcutPath -TargetPath $psExe -Arguments $joinedArgs -WorkingDirectory (Split-Path -Path $scriptFull)

if ($created) {
  Write-Host "Created Startup shortcut: $shortcutPath"
  return
}

# Fallback: create .cmd launcher
$cmdPath = Join-Path $startupFolder ($TaskName + '.cmd')
# For .cmd we use powershell.exe hidden/minimized for compatibility
$psArgs = ('-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "{0}" -IntervalMinutes {1} -Style {2}{3}{4}' -f $scriptFull, $IntervalMinutes, $Style, ($Recurse ? ' -Recurse' : ''), ($LockScreen ? ' -LockScreen' : ''))
$lines = @(
  '@echo off',
  ('start "" /min powershell {0}' -f $psArgs)
)
Set-Content -LiteralPath $cmdPath -Value $lines -Encoding ASCII
Write-Host "Created Startup launcher: $cmdPath"
