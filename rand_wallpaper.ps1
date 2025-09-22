
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
  [string]$ImageFolder = (Join-Path (Join-Path $env:USERPROFILE 'Pictures') 'Wallpapers'),

  [int]$IntervalMinutes = 240,

  [ValidateSet('Fill','Fit','Stretch','Center','Tile','Span')]
  [string]$Style = 'Stretch',

  [switch]$Recurse,

  [switch]$Once,

  [switch]$InstallBackground,

  [switch]$UninstallBackground,

  [string]$TaskName = 'RandomWallpaper',

  [switch]$Tray,

  [switch]$InstallStartup,

  [switch]$UninstallStartup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-DefaultImageFolder {
  param([string]$InputFolder)
  if ($InputFolder -and (Test-Path -LiteralPath $InputFolder)) { return (Get-Item -LiteralPath $InputFolder).FullName }

  # Resolve the user's Pictures folder via Known Folders (handles OneDrive redirection)
  try {
    $knownPictures = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyPictures)
  } catch {
    $knownPictures = $null
  }

  # Preferred default: <Pictures>\Wallpapers (using Known Folder if available)
  if ($knownPictures) {
    $knownWall = Join-Path $knownPictures 'Wallpapers'
    if (Test-Path -LiteralPath $knownWall) { return (Get-Item -LiteralPath $knownWall).FullName }
  }

  # Next, try %USERPROFILE%\Pictures\Wallpapers explicitly
  $pictures = Join-Path $env:USERPROFILE 'Pictures'
  $defaultWall = Join-Path $pictures 'Wallpapers'
  if (Test-Path -LiteralPath $defaultWall) { return (Get-Item -LiteralPath $defaultWall).FullName }

  # Fallbacks: the Pictures folder itself (Known Folder first), then %USERPROFILE%\Pictures, then %USERPROFILE%\wallpapers
  if ($knownPictures -and (Test-Path -LiteralPath $knownPictures)) { return (Get-Item -LiteralPath $knownPictures).FullName }
  if (Test-Path -LiteralPath $pictures) { return (Get-Item -LiteralPath $pictures).FullName }
  $userWall = Join-Path $env:USERPROFILE 'wallpapers'
  if (Test-Path -LiteralPath $userWall) { return (Get-Item -LiteralPath $userWall).FullName }

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

  # SystemParametersInfo has historically worked best with BMP files.
  # Convert non-BMP images to a temporary BMP before applying, to maximize reliability.
  $pathToApply = $ImagePath
  try {
    $ext = [System.IO.Path]::GetExtension($ImagePath)
    if ($ext -and $ext.ToLowerInvariant() -ne '.bmp') {
      Add-Type -AssemblyName System.Drawing -ErrorAction Stop
      $tempDir = Join-Path $env:APPDATA 'RandomWallpaper'
      if (-not (Test-Path -LiteralPath $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
      $bmpPath = Join-Path $tempDir 'current_wallpaper.bmp'
      $img = [System.Drawing.Image]::FromFile($ImagePath)
      try {
        $img.Save($bmpPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
      } finally {
        $img.Dispose()
      }
      $pathToApply = $bmpPath
    }
  } catch {
    # If conversion fails for any reason, fall back to original path.
    $pathToApply = $ImagePath
  }

  # SPI_SETDESKWALLPAPER = 20, SPIF_UPDATEINIFILE = 0x01, SPIF_SENDWININICHANGE = 0x02
  $SPI_SETDESKWALLPAPER = 20
  $flags = 0x01 -bor 0x02
  $ok = [Win32.NativeMethods]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $pathToApply, $flags)
  if (-not $ok) {
    $err = [System.Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw "SystemParametersInfo failed with error code $err for path: $pathToApply"
  }
}

function Test-IsAdmin {
  try {
    $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wp = New-Object Security.Principal.WindowsPrincipal($wi)
    return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}






function Get-ConfigPath {
  $dir = Join-Path $env:APPDATA 'RandomWallpaper'
  if (-not (Test-Path -LiteralPath $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
  Join-Path $dir 'config.json'
}

function Load-Config {
  param(
    [string]$DefaultFolder,
    [int]$DefaultIntervalMinutes,
    [string]$DefaultStyle,
    [bool]$DefaultRecurse,
    [bool]$DefaultScheduleEnabled
  )
  $path = Get-ConfigPath
  if (Test-Path -LiteralPath $path) {
    try { return Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -Depth 5 } catch {}
  }
  [pscustomobject]@{
    ImageFolder     = $DefaultFolder
    IntervalMinutes = $DefaultIntervalMinutes
    Style           = $DefaultStyle
    Recurse         = $DefaultRecurse
    ScheduleEnabled = $DefaultScheduleEnabled
  }
}

function Save-Config {
  param([Parameter(Mandatory)][object]$Config)
  $path = Get-ConfigPath
  $Config | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
}

# Tray icon helpers
$script:generatedIconHandle = [System.IntPtr]::Zero
$script:trayIcon = $null

function New-ImageSymbolIcon {
  param([int]$Width = 16, [int]$Height = 16)
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $bmp = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)
  # Frame
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::DarkSlateGray, 1.8)
  $g.DrawRectangle($pen, 2, 3, $Width-5, $Height-6)
  # Mountain
  $mount = [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.25), [int]($Height*0.85))),
    (New-Object System.Drawing.Point([int]($Width*0.55), [int]($Height*0.50))),
    (New-Object System.Drawing.Point([int]($Width*0.85), [int]($Height*0.85)))
  )
  $g.FillPolygon([System.Drawing.Brushes]::ForestGreen, $mount)
  # Sun
  $g.FillEllipse([System.Drawing.Brushes]::Gold, [int]($Width*0.60), [int]($Height*0.18), [int]($Width*0.25), [int]($Height*0.25))
  $g.Dispose()
  $hicon = $bmp.GetHicon()
  $script:generatedIconHandle = $hicon
  $ico = [System.Drawing.Icon]::FromHandle($hicon)
  $bmp.Dispose()
  return $ico
}

function Save-IconToFile {
  param([Parameter(Mandatory)][System.Drawing.Icon]$Icon, [Parameter(Mandatory)][string]$Path)
  try {
    $dir = Split-Path -LiteralPath $Path -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::Read)
    try { $Icon.Save($fs) } finally { $fs.Dispose() }
  } catch {
    throw "Failed to save icon to '$Path': $($_.Exception.Message)"
  }
}

function New-ImageSymbolBitmap {
  param([int]$Width = 16, [int]$Height = 16)
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $bmp = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)
  # Frame
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::DarkSlateGray, [float]([Math]::Max(1.2, $Width*0.11)))
  $g.DrawRectangle($pen, [int]($Width*0.12), [int]($Height*0.18), [int]($Width*0.76), [int]($Height*0.66))
  # Mountain polygon
  $mount = [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.24), [int]($Height*0.80))),
    (New-Object System.Drawing.Point([int]($Width*0.52), [int]($Height*0.48))),
    (New-Object System.Drawing.Point([int]($Width*0.82), [int]($Height*0.80)))
  )
  $g.FillPolygon([System.Drawing.Brushes]::ForestGreen, $mount)
  # Foreground hill
  $g.FillPolygon([System.Drawing.Brushes]::SeaGreen, [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.18), [int]($Height*0.82))),
    (New-Object System.Drawing.Point([int]($Width*0.35), [int]($Height*0.62))),
    (New-Object System.Drawing.Point([int]($Width*0.60), [int]($Height*0.82)))
  ))
  # Sun
  $sunD = [int]([Math]::Max(3, $Width*0.22))
  $g.FillEllipse([System.Drawing.Brushes]::Gold, [int]($Width*0.60), [int]($Height*0.18), $sunD, $sunD)
  $g.Dispose()
  return $bmp
}

function Write-MultiSizeIco {
  param([Parameter(Mandatory)][string]$Path, [int[]]$Sizes = @(16,24,32,48,64,128,256))
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $images = New-Object System.Collections.Generic.List[object]
  foreach ($s in $Sizes) {
    try {
      $bmp = New-ImageSymbolBitmap -Width $s -Height $s
      $ms = New-Object System.IO.MemoryStream
      $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
      $bmp.Dispose()
      $images.Add([pscustomobject]@{ W = $s; H = $s; Bytes = $ms.ToArray() })
      $ms.Dispose()
    } catch {}
  }
  if ($images.Count -eq 0) { throw 'Failed to generate icon images.' }

  $count = [uint16]$images.Count
  $offset = 6 + (16 * $images.Count)
  $msOut = New-Object System.IO.MemoryStream
  $bw = New-Object System.IO.BinaryWriter($msOut)
  [void]$bw.Write([uint16]0)   # reserved
  [void]$bw.Write([uint16]1)   # ICO type
  [void]$bw.Write([uint16]$count) # image count

  foreach ($img in $images) {
    $wByte = if ($img.W -ge 256) { [byte]0 } else { [byte]$img.W }
    $hByte = if ($img.H -ge 256) { [byte]0 } else { [byte]$img.H }
    [void]$bw.Write([byte]$wByte)            # width
    [void]$bw.Write([byte]$hByte)            # height
    [void]$bw.Write([byte]0)                 # color count
    [void]$bw.Write([byte]0)                 # reserved
    [void]$bw.Write([uint16]0)               # planes (0 for PNG)
    [void]$bw.Write([uint16]32)              # bit count (informational)
    [void]$bw.Write([uint32]$img.Bytes.Length) # bytes in resource
    [void]$bw.Write([uint32]$offset)         # image data offset
    $img | Add-Member -NotePropertyName Offset -NotePropertyValue $offset -Force
    $offset += $img.Bytes.Length
  }

  foreach ($img in $images) { $bw.Write($img.Bytes) }
  $bw.Flush()

  $dir = Split-Path -LiteralPath $Path -Parent
  if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  [System.IO.File]::WriteAllBytes($Path, $msOut.ToArray())
  $bw.Dispose(); $msOut.Dispose()
}

function Ensure-TrayIconFile {
  param([string]$Path = (Join-Path $PSScriptRoot 'tray.ico'))
  $alt = Join-Path $PSScriptRoot 'icon.ico'
  if (Test-Path -LiteralPath $Path -or Test-Path -LiteralPath $alt) { return }
  try {
    Write-MultiSizeIco -Path $Path -Sizes @(16,24,32,48,64,128,256)
  } catch {
    try {
      $ico = New-ImageSymbolIcon -Width 16 -Height 16
      Save-IconToFile -Icon $ico -Path $Path
    } catch {}
    finally {
      if ($ico) { try { $ico.Dispose() } catch {} }
      if ($script:generatedIconHandle -and $script:generatedIconHandle -ne [System.IntPtr]::Zero) { try { [Win32.IconHelper]::DestroyIcon($script:generatedIconHandle) | Out-Null } catch {} }
    }
  }
}

function Get-PreferredTrayIcon {
  try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
    $size = [System.Windows.Forms.SystemInformation]::SmallIconSize
  } catch {
    $size = @{ Width = 16; Height = 16 }
  }
  # Try to ensure a default icon file exists so external consumers (e.g., Startup shortcut) can use it
  try { Ensure-TrayIconFile } catch {}
  $candidates = @(
    (Join-Path $PSScriptRoot 'tray.ico'),
    (Join-Path $PSScriptRoot 'icon.ico')
  )
  foreach ($p in $candidates) {
    if (Test-Path -LiteralPath $p) {
      try { return (New-Object System.Drawing.Icon($p)) } catch {}
    }
  }
  return (New-ImageSymbolIcon -Width $size.Width -Height $size.Height)
}

# P/Invoke to free HICON handles when generated
Add-Type -Namespace Win32 -Name IconHelper -MemberDefinition @"
  [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
  public static extern bool DestroyIcon(System.IntPtr hIcon);
"@ -ErrorAction SilentlyContinue

function Get-StartupShortcutPath {
  $startup = [Environment]::GetFolderPath([Environment+SpecialFolder]::Startup)
  if (-not $startup) { throw 'Failed to resolve Startup folder.' }
  Join-Path $startup 'Random Wallpaper (Tray).lnk'
}

function Test-StartupShortcutExists {
  $lnk = Get-StartupShortcutPath
  Test-Path -LiteralPath $lnk
}

function Install-StartupShortcut {
  param([Parameter(Mandatory)][string]$ScriptPath)
  try {
    $lnk = Get-StartupShortcutPath
  } catch {
    throw "Get-StartupShortcutPath failed: $($_.Exception.Message)"
  }
  try {
    $wsh = New-Object -ComObject WScript.Shell
  } catch {
    throw "Creating WScript.Shell COM object failed: $($_.Exception.Message)"
  }
  try {
    # Prefer Windows PowerShell for hidden STA startup to avoid any console window and ensure STA
    $ps5 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $hostExe = if (Test-Path -LiteralPath $ps5) { $ps5 } else { 'powershell.exe' }
    $args = "-NoProfile -NoLogo -WindowStyle Hidden -Sta -ExecutionPolicy Bypass -File `"$ScriptPath`" -Tray"
  } catch {
    throw "Building arguments failed: $($_.Exception.Message)"
  }
  try {
    $sc = $wsh.CreateShortcut($lnk)
    $sc.TargetPath = $hostExe
    $sc.Arguments = $args
  } catch {
    throw "Creating or configuring .lnk failed: $($_.Exception.Message)"
  }
  try {
    $wd = [System.IO.Path]::GetDirectoryName($ScriptPath)
    if (-not $wd) { $wd = $PSScriptRoot }
    $sc.WorkingDirectory = $wd
  } catch {
    throw "Setting WorkingDirectory failed: $($_.Exception.Message)"
  }
  try {
    $icon1 = Join-Path $PSScriptRoot 'tray.ico'
    $icon2 = Join-Path $PSScriptRoot 'icon.ico'
    if (-not (Test-Path -LiteralPath $icon1) -and -not (Test-Path -LiteralPath $icon2)) {
      try { Ensure-TrayIconFile -Path $icon1 } catch {}
    }
    if (Test-Path -LiteralPath $icon1) { $sc.IconLocation = $icon1 }
    elseif (Test-Path -LiteralPath $icon2) { $sc.IconLocation = $icon2 }
    else { $sc.IconLocation = $hostExe + ',0' }
  } catch {
    throw "Setting icon failed: $($_.Exception.Message)"
  }
  try {
    $sc.Description = 'Random Wallpaper (Tray)'
    $sc.WindowStyle = 7
    $sc.Save()
  } catch {
    throw "Saving shortcut failed: $($_.Exception.Message)"
  }
  return $lnk
}

function Uninstall-StartupShortcut {
  $lnk = Get-StartupShortcutPath
  if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }
}

function Ensure-STA {
  param([string]$ScriptPath, [hashtable]$ForwardParams)
  if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-Sta','-File',("`"{0}`"" -f $ScriptPath),'-Tray')
    foreach ($k in $ForwardParams.Keys) {
      $v = $ForwardParams[$k]
      if ($null -ne $v -and $v -ne '') {
        if ($v -is [switch] -or $v -is [bool]) {
          if ([bool]$v) { $argList += @("-$k") }
        } else {
          $argList += @("-$k", ("`"{0}`"" -f $v))
        }
      }
    }
    # Use Windows PowerShell for guaranteed -STA support on Windows
    $ps5 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $exe = if (Test-Path -LiteralPath $ps5) { $ps5 } else { 'powershell.exe' }
    Start-Process $exe -ArgumentList ($argList -join ' ') -WindowStyle Hidden | Out-Null
    return $false
  }
  return $true
}

function Start-TrayApp {
  param(
    [string]$InitialImageFolder,
    [int]$IntervalMinutes,
    [string]$Style,
    [switch]$Recurse
  )
  # Determine script path for re-launching and shortcuts:
  # - Use $PSCommandPath (available within functions and compatible with Set-StrictMode).
  # - Avoid $MyInvocation.MyCommand.Path here: inside a function it is a FunctionInfo without a Path property,
  #   which triggers a property access error under Set-StrictMode.
  # - Fallback to $PSScriptRoot if $PSCommandPath isn't populated (older hosts or dot-sourced execution).
  $scriptPath = $PSCommandPath
  if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'rand_wallpaper.ps1' }
  $forward = @{ ImageFolder = $InitialImageFolder; IntervalMinutes = $IntervalMinutes; Style = $Style; Recurse = [bool]$Recurse }
  if (-not (Ensure-STA -ScriptPath $scriptPath -ForwardParams $forward)) { return }

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}

  # Ensure a tray icon file exists if none has been provided
  try { Ensure-TrayIconFile } catch {}

  # Resolve defaults
  $resolvedFolder = Resolve-DefaultImageFolder -InputFolder $InitialImageFolder
  $useRecurse = if ($PSBoundParameters.ContainsKey('Recurse')) { [bool]$Recurse } else { $true }
  $cfg = Load-Config -DefaultFolder $resolvedFolder -DefaultIntervalMinutes $IntervalMinutes -DefaultStyle $Style -DefaultRecurse $useRecurse -DefaultScheduleEnabled $true

  # UI objects
  $context = New-Object System.Windows.Forms.ApplicationContext
  $notify  = New-Object System.Windows.Forms.NotifyIcon
  $script:trayIcon = Get-PreferredTrayIcon
  $notify.Icon = $script:trayIcon
  $notify.Visible = $true
  $notify.Text = "Random Wallpaper"

  $menu = New-Object System.Windows.Forms.ContextMenuStrip
  $itemChangeNow    = $menu.Items.Add("Change now")
  $itemSetFolder    = $menu.Items.Add("Set folder...")
  $itemSetInterval  = $menu.Items.Add("Set interval (minutes)...")
  $itemIncludeSub   = New-Object System.Windows.Forms.ToolStripMenuItem("Include subfolders")
  $itemIncludeSub.CheckOnClick = $true
  $null = $menu.Items.Add($itemIncludeSub)
  $itemStyle        = New-Object System.Windows.Forms.ToolStripMenuItem("Style")
  $styles = @('Fill','Fit','Stretch','Center','Tile','Span')
  foreach ($s in $styles) {
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem($s)
    $mi.CheckOnClick = $true
    if ($s -eq $cfg.Style) { $mi.Checked = $true }
    $mi.Add_Click({
      param($sender,$args)
      foreach ($x in $itemStyle.DropDownItems) { if ($x -ne $sender) { $x.Checked = $false } }
      $script:cfg.Style = $sender.Text
      Save-Config -Config $script:cfg
    })
    $null = $itemStyle.DropDownItems.Add($mi)
  }
  $null = $menu.Items.Add($itemStyle)
  $itemSchedule      = New-Object System.Windows.Forms.ToolStripMenuItem("Scheduling enabled")
  $itemSchedule.CheckOnClick = $true
  $null = $menu.Items.Add($itemSchedule)
  $itemStartup       = New-Object System.Windows.Forms.ToolStripMenuItem("Start at logon")
  $itemStartup.CheckOnClick = $true
  $null = $menu.Items.Add($itemStartup)
  $menu.Items.Add('-') | Out-Null
  $itemOpenFolder    = $menu.Items.Add("Open images folder")
  $menu.Items.Add('-') | Out-Null
  $itemExit          = $menu.Items.Add("Exit")

  $notify.ContextMenuStrip = $menu

  # Timer
  $timer = New-Object System.Windows.Forms.Timer
  function Update-Timer {
    $timer.Interval = [Math]::Max(1, [int]$script:cfg.IntervalMinutes) * 60000
    $timer.Enabled = [bool]$script:cfg.ScheduleEnabled
    $itemSchedule.Checked = $timer.Enabled
  }
  $script:cfg = $cfg
  $itemIncludeSub.Checked = [bool]$script:cfg.Recurse
  Update-Timer
  $itemStartup.Checked = (Test-StartupShortcutExists)

  $timer.add_Tick({
    try {
      Set-RandomWallpaperOnce -Folder $script:cfg.ImageFolder -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style
    } catch {
      [System.Windows.Forms.MessageBox]::Show("Failed to set wallpaper: $($_.Exception.Message)","Random Wallpaper",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
  })

  # Handlers
  $itemChangeNow.Add_Click({
    try { Set-RandomWallpaperOnce -Folder $script:cfg.ImageFolder -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style } catch {}
  })
  $notify.Add_DoubleClick({ $itemChangeNow.PerformClick() })

  $itemSetFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $script:cfg.ImageFolder
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $script:cfg.ImageFolder = $dlg.SelectedPath
      Save-Config -Config $script:cfg
    }
  })

  $itemSetInterval.Add_Click({
    $current = [string]$script:cfg.IntervalMinutes
    try {
      $input = [Microsoft.VisualBasic.Interaction]::InputBox("Change every N minutes:","Random Wallpaper",$current)
    } catch {
      $form = New-Object System.Windows.Forms.Form
      $form.Text = "Random Wallpaper"
      $form.Width = 320; $form.Height = 140
      $tb = New-Object System.Windows.Forms.TextBox
      $tb.Text = $current; $tb.Left = 12; $tb.Top = 12; $tb.Width = 280
      $ok = New-Object System.Windows.Forms.Button
      $ok.Text = "OK"; $ok.Left = 140; $ok.Top = 50; $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
      $cancel = New-Object System.Windows.Forms.Button
      $cancel.Text = "Cancel"; $cancel.Left = 220; $cancel.Top = 50; $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
      $form.Controls.AddRange(@($tb,$ok,$cancel))
      $form.AcceptButton = $ok; $form.CancelButton = $cancel
      $input = if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text } else { "" }
      $form.Dispose()
    }
    if ($input -and $input.Trim().Length -gt 0) {
      $parsed = 0
      if ([int]::TryParse($input, [ref]$parsed)) {
        $script:cfg.IntervalMinutes = [int]$parsed
        Save-Config -Config $script:cfg
        Update-Timer
      } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid integer number of minutes.","Random Wallpaper") | Out-Null
      }
    }
  })

  $itemIncludeSub.Add_Click({
    $script:cfg.Recurse = $itemIncludeSub.Checked
    Save-Config -Config $script:cfg
  })

  $itemSchedule.Add_Click({
    $script:cfg.ScheduleEnabled = $itemSchedule.Checked
    Save-Config -Config $script:cfg
    Update-Timer
  })

  $itemStartup.Add_Click({
    try {
      if ($itemStartup.Checked) { Install-StartupShortcut -ScriptPath $scriptPath }
      else { Uninstall-StartupShortcut }
    } catch {
      $itemStartup.Checked = -not $itemStartup.Checked
      [System.Windows.Forms.MessageBox]::Show("Failed to update startup shortcut: $($_.Exception.Message)","Random Wallpaper") | Out-Null
    }
  })

  $itemOpenFolder.Add_Click({
    try { Start-Process explorer.exe -ArgumentList ("`"{0}`"" -f $script:cfg.ImageFolder) } catch {}
  })

  $itemExit.Add_Click({
    $timer.Stop()
    $notify.Visible = $false
    $notify.Dispose()
    if ($script:trayIcon) { try { $script:trayIcon.Dispose() } catch {} }
    if ($script:generatedIconHandle -and $script:generatedIconHandle -ne [System.IntPtr]::Zero) { try { [Win32.IconHelper]::DestroyIcon($script:generatedIconHandle) | Out-Null } catch {} }
    $context.ExitThread()
  })

  # Apply one immediately on start
  try { Set-RandomWallpaperOnce -Folder $script:cfg.ImageFolder -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style } catch {}

  [System.Windows.Forms.Application]::Run($context)
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

  if (-not (Test-IsAdmin)) {
    Write-Host "Attempting to re-launch with administrator privileges to uninstall Scheduled Task..."
    Start-Process pwsh -Verb RunAs -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "$PSScriptRoot\rand_wallpaper.ps1", "-UninstallBackground", "-TaskName", "'$TaskName'"
    return
  }
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
# Resolve the on-disk script path for scheduled tasks and startup shortcuts:
# - Prefer $PSCommandPath for reliability in both script and function scope.
# - Avoid $MyInvocation.MyCommand.Path here to prevent StrictMode property errors in certain contexts.
# - Fallback to $PSScriptRoot when $PSCommandPath is not set (older hosts or when dot-sourced).
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'rand_wallpaper.ps1' }
if ($UninstallBackground) {
  Uninstall-BackgroundTask -TaskName $TaskName
  return
}
if ($InstallBackground) {
  $resolvedFolderForTask = Resolve-DefaultImageFolder -InputFolder $ImageFolder
  Install-BackgroundTask -ScriptPath $scriptPath -ResolvedImageFolder $resolvedFolderForTask -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse -TaskName $TaskName
  return
}

# Startup shortcut install/uninstall
if ($InstallStartup) {
  Install-StartupShortcut -ScriptPath $scriptPath
  return
}
if ($UninstallStartup) {
  Uninstall-StartupShortcut
  return
}

# If tray mode requested, start the tray app (Windows Forms) and exit this instance
if ($Tray) {
  Start-TrayApp -InitialImageFolder $ImageFolder -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse
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
