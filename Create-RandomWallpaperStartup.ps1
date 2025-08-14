<#
.SYNOPSIS
  Creates a Startup entry that launches rand_wallpaper.ps1 at user logon with specified options.

.DESCRIPTION
  - Tries to create a .lnk shortcut in the user's Startup folder using WScript.Shell
  - Falls back to creating a .cmd launcher if COM is unavailable
  - Prefers Windows PowerShell (powershell.exe) for better hidden window support, falls back to pwsh.exe if needed
  - Passes through Interval/Style/Recurse options to the wallpaper script (relies on script's default ImageFolder)

.PARAMETER ScriptPath
  Path to rand_wallpaper.ps1. Defaults to rand_wallpaper.ps1 in the same directory as this script.

.PARAMETER IntervalMinutes
  Interval in minutes between changes. Default 15.

.PARAMETER Style
  Fill, Fit, Stretch, Center, Tile, Span. Default Fill.

.PARAMETER Recurse
  Include subfolders when picking images.

.PARAMETER TaskName
  Used to name files created; default "RandomWallpaper"

.EXAMPLE
  # Create Startup entry using defaults
  ./Create-RandomWallpaperStartup.ps1 -IntervalMinutes 15 -Style Fill -Recurse
#>
param(
  [string]$ScriptPath = (Join-Path $PSScriptRoot 'rand_wallpaper.ps1'),
  [int]$IntervalMinutes = 15,
  [ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')]
  [string]$Style = 'Fill',
  [switch]$Recurse,
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
    [Parameter(Mandatory)][string]$PsExe,
    [Parameter(Mandatory)][string]$ScriptPath,
    [Parameter(Mandatory)][string]$Args
  )
  # Build a small .cmd that starts the chosen PowerShell host minimized with the given args
  $lines = @(
    '@echo off',
    ('start "" /min "{0}" {1}' -f $PsExe, $Args)
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
if ($Recurse)          { $argList += '-Recurse' }
$joinedArgs = ($argList -join ' ')

# Try .lnk first
$shortcutPath = Join-Path $startupFolder ($TaskName + '.lnk')
$created = New-StartupShortcut -ShortcutPath $shortcutPath -TargetPath $psExe -Arguments $joinedArgs -WorkingDirectory (Split-Path -Path $scriptFull)

if ($created) {
  Write-Host "Created Startup shortcut: $shortcutPath"
  return
}

# Fallback: create .cmd launcher using the same PowerShell host and arguments
$cmdPath = Join-Path $startupFolder ($TaskName + '.cmd')
$psArgs = $joinedArgs
New-StartupCmdLauncher -CmdPath $cmdPath -PsExe $psExe -ScriptPath $scriptFull -Args $psArgs
Write-Host "Created Startup launcher: $cmdPath"
