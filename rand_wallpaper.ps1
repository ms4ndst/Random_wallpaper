
<#
.SYNOPSIS
  Set a random desktop wallpaper from a folder. Immediately sets on start, then optionally loops on an interval.

.DESCRIPTION
  - Selects a random image from one or more configured folders (optionally including subfolders)
  - Sets registry values for wallpaper style (Fill, Fit, Stretch, Center, Tile, Span)
  - Applies the wallpaper using SystemParametersInfo (works on Windows 10/11)
  - Immediately sets a new wallpaper when the script begins
  - Optionally runs continuously and changes the wallpaper on a timer
  - Avoids repeating recently shown wallpapers (history-based dedup)
  - Supports multiple source folders and time-of-day folder switching

.PARAMETER ImageFolder
  Primary folder containing images. Defaults to "Pictures\Wallpapers" under the current user profile if it exists,
  otherwise falls back to the user's Pictures folder.

.PARAMETER IntervalMinutes
  Minutes between wallpaper changes when running continuously. Default 240.

.PARAMETER Style
  Wallpaper style: Fill, Fit, Stretch, Center, Tile, Span. Default Stretch.

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
  Requires pwsh/PowerShell 7+ on Windows 10/11. Supported image types: jpg, jpeg, png, bmp, webp.
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

# History list for avoiding recently shown wallpapers (in-memory, per session)
$script:wallpaperHistory = [System.Collections.Generic.List[string]]::new()

function Resolve-DefaultImageFolder {
  param([string]$InputFolder)
  if ($InputFolder -and (Test-Path -LiteralPath $InputFolder)) { return (Get-Item -LiteralPath $InputFolder).FullName }

  try {
    $knownPictures = [Environment]::GetFolderPath([Environment+SpecialFolder]::MyPictures)
  } catch {
    $knownPictures = $null
  }

  if ($knownPictures) {
    $knownWall = Join-Path $knownPictures 'Wallpapers'
    if (Test-Path -LiteralPath $knownWall) { return (Get-Item -LiteralPath $knownWall).FullName }
  }

  $pictures = Join-Path $env:USERPROFILE 'Pictures'
  $defaultWall = Join-Path $pictures 'Wallpapers'
  if (Test-Path -LiteralPath $defaultWall) { return (Get-Item -LiteralPath $defaultWall).FullName }

  if ($knownPictures -and (Test-Path -LiteralPath $knownPictures)) { return (Get-Item -LiteralPath $knownPictures).FullName }
  if (Test-Path -LiteralPath $pictures) { return (Get-Item -LiteralPath $pictures).FullName }
  $userWall = Join-Path $env:USERPROFILE 'wallpapers'
  if (Test-Path -LiteralPath $userWall) { return (Get-Item -LiteralPath $userWall).FullName }

  throw "No valid image folder found. Provide -ImageFolder with a valid path."
}

# Returns the effective folder list considering time-of-day overrides and extra folders.
function Get-EffectiveFolders {
  param([Parameter(Mandatory)][object]$Config)
  $hour = (Get-Date).Hour
  $isDaytime = ($hour -ge 6 -and $hour -lt 20)

  if ($isDaytime -and
      $Config.PSObject.Properties['DayFolder'] -and
      $Config.DayFolder -and
      (Test-Path -LiteralPath $Config.DayFolder)) {
    return @($Config.DayFolder)
  }
  if (-not $isDaytime -and
      $Config.PSObject.Properties['NightFolder'] -and
      $Config.NightFolder -and
      (Test-Path -LiteralPath $Config.NightFolder)) {
    return @($Config.NightFolder)
  }

  $result = [System.Collections.Generic.List[string]]::new()
  if ($Config.ImageFolder) { $result.Add($Config.ImageFolder) }
  if ($Config.PSObject.Properties['ExtraFolders'] -and $Config.ExtraFolders) {
    foreach ($f in $Config.ExtraFolders) {
      if ($f -and (Test-Path -LiteralPath $f) -and -not $result.Contains($f)) {
        $result.Add($f)
      }
    }
  }
  if ($result.Count -eq 0) { throw "No valid image folders configured." }
  return $result.ToArray()
}

function Get-RandomImagePath {
  param(
    [Parameter(Mandatory)] [string[]]$Folders,
    [switch]$IncludeSubfolders
  )
  $patterns = @('*.jpg','*.jpeg','*.png','*.bmp','*.webp')
  $fileList = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
  foreach ($folder in $Folders) {
    if (-not (Test-Path -LiteralPath $folder)) { continue }
    foreach ($pattern in $patterns) {
      $found = Get-ChildItem -LiteralPath $folder -Filter $pattern -File -ErrorAction SilentlyContinue -Recurse:$IncludeSubfolders
      if ($found) { foreach ($f in $found) { $fileList.Add($f) } }
    }
  }

  # Deduplicate by full path (case-insensitive) to avoid bias from overlapping patterns
  $files = @($fileList | Sort-Object { $_.FullName.ToLowerInvariant() } -Unique)

  if ($files.Count -eq 0) {
    $folderList = $Folders -join "', '"
    throw "No images found in '$folderList'$(if($IncludeSubfolders){' (including subfolders)'}). Supported: jpg, jpeg, png, bmp, webp."
  }

  # Filter out recently shown images to prevent immediate repeats
  if ($script:wallpaperHistory.Count -gt 0 -and $script:wallpaperHistory.Count -lt $files.Count) {
    $histSet = [System.Collections.Generic.HashSet[string]]::new(
      [string[]]($script:wallpaperHistory | ForEach-Object { $_.ToLowerInvariant() }),
      [System.StringComparer]::OrdinalIgnoreCase
    )
    $candidates = @($files | Where-Object { -not $histSet.Contains($_.FullName) })
    if ($candidates.Count -gt 0) { $files = $candidates }
  }

  $chosen = ($files | Get-Random).FullName

  # Update history; cap at half the total collection size or 20, whichever is larger
  $script:wallpaperHistory.Add($chosen)
  $maxHist = [Math]::Max(20, [int]($fileList.Count * 0.5))
  while ($script:wallpaperHistory.Count -gt $maxHist) { $script:wallpaperHistory.RemoveAt(0) }

  return $chosen
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
    'Span'    { $wallpaperStyle = '22'; $tile = '0' }
  }
  Set-ItemProperty -Path $desktopKey -Name WallpaperStyle -Value $wallpaperStyle -Type String | Out-Null
  Set-ItemProperty -Path $desktopKey -Name TileWallpaper  -Value $tile           -Type String | Out-Null
}

Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @"
  [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
  public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
"@

function Apply-Wallpaper {
  param([Parameter(Mandatory)][string]$ImagePath)

  $pathToApply = $ImagePath
  try {
    $ext = [System.IO.Path]::GetExtension($ImagePath).ToLowerInvariant()
    if ($ext -and $ext -ne '.bmp') {
      $tempDir = Join-Path $env:APPDATA 'RandomWallpaper'
      if (-not (Test-Path -LiteralPath $tempDir)) { New-Item -Path $tempDir -ItemType Directory -Force | Out-Null }
      $bmpPath = Join-Path $tempDir 'current_wallpaper.bmp'

      if ($ext -eq '.webp') {
        # webp requires WPF BitmapDecoder (System.Drawing does not support it)
        try {
          Add-Type -AssemblyName PresentationCore -ErrorAction Stop
          $stream = [System.IO.File]::OpenRead($ImagePath)
          try {
            $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
              $stream,
              [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
              [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            )
            $frame = $decoder.Frames[0]
            $encoder = New-Object System.Windows.Media.Imaging.BmpBitmapEncoder
            $encoder.Frames.Add($frame)
            $outStream = [System.IO.File]::Open($bmpPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
            try { $encoder.Save($outStream) } finally { $outStream.Dispose() }
          } finally { $stream.Dispose() }
          $pathToApply = $bmpPath
        } catch {
          $pathToApply = $ImagePath  # fallback: try raw path
        }
      } else {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $img = [System.Drawing.Image]::FromFile($ImagePath)
        try {
          $img.Save($bmpPath, [System.Drawing.Imaging.ImageFormat]::Bmp)
        } finally {
          $img.Dispose()
        }
        $pathToApply = $bmpPath
      }
    }
  } catch {
    $pathToApply = $ImagePath
  }

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
  $cfg = $null
  if (Test-Path -LiteralPath $path) {
    try { $cfg = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -Depth 5 } catch {}
  }
  if (-not $cfg) {
    $cfg = [pscustomobject]@{
      ImageFolder      = $DefaultFolder
      ExtraFolders     = @()
      IntervalMinutes  = $DefaultIntervalMinutes
      Style            = $DefaultStyle
      Recurse          = $DefaultRecurse
      ScheduleEnabled  = $DefaultScheduleEnabled
      DayFolder        = $null
      NightFolder      = $null
    }
  } else {
    # Migrate/add fields absent in older config versions
    if (-not $cfg.PSObject.Properties['ExtraFolders'])  { $cfg | Add-Member -NotePropertyName ExtraFolders  -NotePropertyValue @()   }
    if (-not $cfg.PSObject.Properties['DayFolder'])     { $cfg | Add-Member -NotePropertyName DayFolder     -NotePropertyValue $null  }
    if (-not $cfg.PSObject.Properties['NightFolder'])   { $cfg | Add-Member -NotePropertyName NightFolder   -NotePropertyValue $null  }
  }
  return $cfg
}

function Save-Config {
  param([Parameter(Mandatory)][object]$Config)
  $path = Get-ConfigPath
  $Config | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $path -Encoding UTF8
}

# ── Tray icon helpers ──────────────────────────────────────────────────────────
$script:generatedIconHandle = [System.IntPtr]::Zero
$script:trayIcon = $null

function New-ImageSymbolIcon {
  param([int]$Width = 16, [int]$Height = 16)
  Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue
  $bmp = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($bmp)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
  $g.Clear([System.Drawing.Color]::Transparent)
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::DarkSlateGray, 1.8)
  $g.DrawRectangle($pen, 2, 3, $Width-5, $Height-6)
  $mount = [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.25), [int]($Height*0.85))),
    (New-Object System.Drawing.Point([int]($Width*0.55), [int]($Height*0.50))),
    (New-Object System.Drawing.Point([int]($Width*0.85), [int]($Height*0.85)))
  )
  $g.FillPolygon([System.Drawing.Brushes]::ForestGreen, $mount)
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
  $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::DarkSlateGray, [float]([Math]::Max(1.2, $Width*0.11)))
  $g.DrawRectangle($pen, [int]($Width*0.12), [int]($Height*0.18), [int]($Width*0.76), [int]($Height*0.66))
  $mount = [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.24), [int]($Height*0.80))),
    (New-Object System.Drawing.Point([int]($Width*0.52), [int]($Height*0.48))),
    (New-Object System.Drawing.Point([int]($Width*0.82), [int]($Height*0.80)))
  )
  $g.FillPolygon([System.Drawing.Brushes]::ForestGreen, $mount)
  $g.FillPolygon([System.Drawing.Brushes]::SeaGreen, [System.Drawing.Point[]]@(
    (New-Object System.Drawing.Point([int]($Width*0.18), [int]($Height*0.82))),
    (New-Object System.Drawing.Point([int]($Width*0.35), [int]($Height*0.62))),
    (New-Object System.Drawing.Point([int]($Width*0.60), [int]($Height*0.82)))
  ))
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

  $count  = [uint16]$images.Count
  $offset = 6 + (16 * $images.Count)
  $msOut  = New-Object System.IO.MemoryStream
  $bw     = New-Object System.IO.BinaryWriter($msOut)
  [void]$bw.Write([uint16]0)
  [void]$bw.Write([uint16]1)
  [void]$bw.Write([uint16]$count)

  foreach ($img in $images) {
    $wByte = if ($img.W -ge 256) { [byte]0 } else { [byte]$img.W }
    $hByte = if ($img.H -ge 256) { [byte]0 } else { [byte]$img.H }
    [void]$bw.Write([byte]$wByte)
    [void]$bw.Write([byte]$hByte)
    [void]$bw.Write([byte]0)
    [void]$bw.Write([byte]0)
    [void]$bw.Write([uint16]0)
    [void]$bw.Write([uint16]32)
    [void]$bw.Write([uint32]$img.Bytes.Length)
    [void]$bw.Write([uint32]$offset)
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
      if ($script:generatedIconHandle -and $script:generatedIconHandle -ne [System.IntPtr]::Zero) {
        try { [Win32.IconHelper]::DestroyIcon($script:generatedIconHandle) | Out-Null } catch {}
      }
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
  try { $lnk = Get-StartupShortcutPath } catch { throw "Get-StartupShortcutPath failed: $($_.Exception.Message)" }
  try { $wsh = New-Object -ComObject WScript.Shell } catch { throw "Creating WScript.Shell COM object failed: $($_.Exception.Message)" }
  try {
    $pwsh = (Get-Command pwsh -ErrorAction SilentlyContinue)?.Source
    if (-not $pwsh) { $pwsh = 'pwsh.exe' }
    $hostExe = $pwsh
    $args = "-NoProfile -NoLogo -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" -Tray"
  } catch { throw "Building arguments failed: $($_.Exception.Message)" }
  try {
    $sc = $wsh.CreateShortcut($lnk)
    $sc.TargetPath = $hostExe
    $sc.Arguments = $args
  } catch { throw "Creating or configuring .lnk failed: $($_.Exception.Message)" }
  try {
    $wd = [System.IO.Path]::GetDirectoryName($ScriptPath)
    if (-not $wd) { $wd = $PSScriptRoot }
    $sc.WorkingDirectory = $wd
  } catch { throw "Setting WorkingDirectory failed: $($_.Exception.Message)" }
  try {
    $icon1 = Join-Path $PSScriptRoot 'tray.ico'
    $icon2 = Join-Path $PSScriptRoot 'icon.ico'
    if (-not (Test-Path -LiteralPath $icon1) -and -not (Test-Path -LiteralPath $icon2)) {
      try { Ensure-TrayIconFile -Path $icon1 } catch {}
    }
    if (Test-Path -LiteralPath $icon1)      { $sc.IconLocation = $icon1 }
    elseif (Test-Path -LiteralPath $icon2)  { $sc.IconLocation = $icon2 }
    else                                    { $sc.IconLocation = $hostExe + ',0' }
  } catch { throw "Setting icon failed: $($_.Exception.Message)" }
  try {
    $sc.Description = 'Random Wallpaper (Tray)'
    $sc.WindowStyle = 7
    $sc.Save()
  } catch { throw "Saving shortcut failed: $($_.Exception.Message)" }
  return $lnk
}

function Uninstall-StartupShortcut {
  $lnk = Get-StartupShortcutPath
  if (Test-Path -LiteralPath $lnk) { Remove-Item -LiteralPath $lnk -Force }
}

function Ensure-STA {
  param([string]$ScriptPath, [hashtable]$ForwardParams)
  if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',("\`"{0}\`"" -f $ScriptPath),'-Tray')
    foreach ($k in $ForwardParams.Keys) {
      $v = $ForwardParams[$k]
      if ($null -ne $v -and $v -ne '') {
        if ($v -is [switch] -or $v -is [bool]) {
          if ([bool]$v) { $argList += @("-$k") }
        } else {
          $argList += @("-$k", ("\`"{0}\`"" -f $v))
        }
      }
    }
    $exe = (Get-Command pwsh -ErrorAction SilentlyContinue)?.Source
    if (-not $exe) { $exe = 'pwsh.exe' }
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
  $scriptPath = $PSCommandPath
  if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'rand_wallpaper.ps1' }
  $forward = @{ ImageFolder = $InitialImageFolder; IntervalMinutes = $IntervalMinutes; Style = $Style; Recurse = [bool]$Recurse }
  if (-not (Ensure-STA -ScriptPath $scriptPath -ForwardParams $forward)) { return }

  Add-Type -AssemblyName System.Windows.Forms
  Add-Type -AssemblyName System.Drawing
  try { Add-Type -AssemblyName Microsoft.VisualBasic } catch {}

  try { Ensure-TrayIconFile } catch {}

  $resolvedFolder = Resolve-DefaultImageFolder -InputFolder $InitialImageFolder
  $useRecurse = if ($PSBoundParameters.ContainsKey('Recurse')) { [bool]$Recurse } else { $true }
  $cfg = Load-Config -DefaultFolder $resolvedFolder -DefaultIntervalMinutes $IntervalMinutes -DefaultStyle $Style -DefaultRecurse $useRecurse -DefaultScheduleEnabled $true

  $context = New-Object System.Windows.Forms.ApplicationContext
  $notify  = New-Object System.Windows.Forms.NotifyIcon
  $script:trayIcon = Get-PreferredTrayIcon
  $notify.Icon    = $script:trayIcon
  $notify.Visible = $true
  $notify.Text    = "Random Wallpaper"

  $menu             = New-Object System.Windows.Forms.ContextMenuStrip
  $itemChangeNow    = $menu.Items.Add("Change now")
  $itemSelectWall   = $menu.Items.Add("Select wallpaper...")

  # Folders submenu
  $itemFolders      = New-Object System.Windows.Forms.ToolStripMenuItem("Folders")
  $itemSetFolder    = New-Object System.Windows.Forms.ToolStripMenuItem("Set primary folder...")
  $itemAddFolder    = New-Object System.Windows.Forms.ToolStripMenuItem("Add extra folder...")
  $itemClearExtra   = New-Object System.Windows.Forms.ToolStripMenuItem("Clear extra folders")
  $null = $itemFolders.DropDownItems.Add($itemSetFolder)
  $null = $itemFolders.DropDownItems.Add($itemAddFolder)
  $null = $itemFolders.DropDownItems.Add($itemClearExtra)
  $null = $menu.Items.Add($itemFolders)

  $itemSetInterval  = $menu.Items.Add("Set interval (minutes)...")
  $itemIncludeSub   = New-Object System.Windows.Forms.ToolStripMenuItem("Include subfolders")
  $itemIncludeSub.CheckOnClick = $true
  $null = $menu.Items.Add($itemIncludeSub)

  $itemStyle = New-Object System.Windows.Forms.ToolStripMenuItem("Style")
  foreach ($s in @('Fill','Fit','Stretch','Center','Tile','Span')) {
    $mi = New-Object System.Windows.Forms.ToolStripMenuItem($s)
    $mi.CheckOnClick = $true
    if ($s -eq $cfg.Style) { $mi.Checked = $true }
    $mi.Add_Click({
      param($sender,$e)
      foreach ($x in $itemStyle.DropDownItems) { if ($x -ne $sender) { $x.Checked = $false } }
      $script:cfg.Style = $sender.Text
      Save-Config -Config $script:cfg
    })
    $null = $itemStyle.DropDownItems.Add($mi)
  }
  $null = $menu.Items.Add($itemStyle)

  # Time-of-day submenu
  $itemTimeOfDay      = New-Object System.Windows.Forms.ToolStripMenuItem("Time-of-day folders")
  $itemSetDayFolder   = New-Object System.Windows.Forms.ToolStripMenuItem("Set day folder (6am-8pm)...")
  $itemSetNightFolder = New-Object System.Windows.Forms.ToolStripMenuItem("Set night folder (8pm-6am)...")
  $itemClearTOD       = New-Object System.Windows.Forms.ToolStripMenuItem("Clear time-of-day folders")
  $null = $itemTimeOfDay.DropDownItems.Add($itemSetDayFolder)
  $null = $itemTimeOfDay.DropDownItems.Add($itemSetNightFolder)
  $null = $itemTimeOfDay.DropDownItems.Add($itemClearTOD)
  $null = $menu.Items.Add($itemTimeOfDay)

  $itemSchedule = New-Object System.Windows.Forms.ToolStripMenuItem("Scheduling enabled")
  $itemSchedule.CheckOnClick = $true
  $null = $menu.Items.Add($itemSchedule)

  $itemStartup = New-Object System.Windows.Forms.ToolStripMenuItem("Start at logon")
  $itemStartup.CheckOnClick = $true
  $null = $menu.Items.Add($itemStartup)

  $menu.Items.Add('-') | Out-Null
  $itemOpenFolder = $menu.Items.Add("Open images folder")
  $menu.Items.Add('-') | Out-Null
  $itemExit = $menu.Items.Add("Exit")

  $notify.ContextMenuStrip = $menu

  $timer = New-Object System.Windows.Forms.Timer
  function Update-Timer {
    $timer.Interval = [Math]::Max(1, [int]$script:cfg.IntervalMinutes) * 60000
    $timer.Enabled  = [bool]$script:cfg.ScheduleEnabled
    $itemSchedule.Checked = $timer.Enabled
  }

  $script:cfg = $cfg
  $itemIncludeSub.Checked = [bool]$script:cfg.Recurse
  Update-Timer
  $itemStartup.Checked = (Test-StartupShortcutExists)

  $timer.add_Tick({
    try {
      $folders = Get-EffectiveFolders -Config $script:cfg
      Set-RandomWallpaperOnce -Folders $folders -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style
    } catch {
      [System.Windows.Forms.MessageBox]::Show(
        "Failed to set wallpaper: $($_.Exception.Message)",
        "Random Wallpaper",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
  })

  $itemChangeNow.Add_Click({
    try {
      $folders = Get-EffectiveFolders -Config $script:cfg
      Set-RandomWallpaperOnce -Folders $folders -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style
    } catch {}
  })
  $notify.Add_DoubleClick({ $itemChangeNow.PerformClick() })

  $itemSelectWall.Add_Click({
    $patterns   = @('*.jpg','*.jpeg','*.png','*.bmp','*.webp')
    $activeFolders = try { Get-EffectiveFolders -Config $script:cfg } catch { @($script:cfg.ImageFolder) }
    $rawFiles   = [System.Collections.Generic.List[System.IO.FileInfo]]::new()
    foreach ($folder in $activeFolders) {
      if (-not (Test-Path -LiteralPath $folder)) { continue }
      foreach ($p in $patterns) {
        $found = Get-ChildItem -LiteralPath $folder -Filter $p -File `
          -Recurse:([bool]$script:cfg.Recurse) -ErrorAction SilentlyContinue
        if ($found) { foreach ($f in $found) { $rawFiles.Add($f) } }
      }
    }
    $imageFiles = @($rawFiles | Sort-Object { $_.FullName.ToLowerInvariant() } -Unique)

    if ($imageFiles.Count -eq 0) {
      [System.Windows.Forms.MessageBox]::Show(
        "No images found in the configured folder(s).",
        "Random Wallpaper",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
      return
    }

    $thumbSize = 180
    $thumbPad  = 10
    $cols      = 4

    $pickerForm = New-Object System.Windows.Forms.Form
    $pickerForm.Text            = "Select Wallpaper  ($($imageFiles.Count) images)"
    $pickerForm.BackColor       = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $pickerForm.ForeColor       = [System.Drawing.Color]::White
    $pickerForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable
    $pickerForm.StartPosition   = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $pickerForm.Width           = $cols * ($thumbSize + $thumbPad * 2) + 40
    $pickerForm.Height          = 680
    $pickerForm.MinimumSize     = New-Object System.Drawing.Size(400, 400)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Dock        = [System.Windows.Forms.DockStyle]::Top
    $searchBox.BackColor   = [System.Drawing.Color]::FromArgb(50, 50, 50)
    $searchBox.ForeColor   = [System.Drawing.Color]::White
    $searchBox.Font        = New-Object System.Drawing.Font('Segoe UI', 10)
    $searchBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $searchBox.Height      = 28

    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.Dock       = [System.Windows.Forms.DockStyle]::Fill
    $flow.AutoScroll = $true
    $flow.BackColor  = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $flow.Padding    = New-Object System.Windows.Forms.Padding(6)

    $pickerForm.Controls.Add($flow)
    $pickerForm.Controls.Add($searchBox)

    $disposeAll = {
      foreach ($tile in $allTiles) {
        foreach ($child in $tile.Controls) {
          if ($child -is [System.Windows.Forms.PictureBox]) {
            try { $child.CancelAsync() } catch {}
            if ($child.Image) {
              try { $child.Image.Dispose() } catch {}
              $child.Image = $null
            }
          }
        }
      }
    }

    # Build tile without loading image (lazy-loaded later)
    $buildTile = {
      param([System.IO.FileInfo]$file)
      $tile = New-Object System.Windows.Forms.Panel
      $tile.Width    = $thumbSize + $thumbPad * 2
      $tile.Height   = $thumbSize + $thumbPad * 2 + 22
      $tile.Margin   = New-Object System.Windows.Forms.Padding(4)
      $tile.BackColor= [System.Drawing.Color]::FromArgb(45, 45, 45)
      $tile.Cursor   = [System.Windows.Forms.Cursors]::Hand
      $tile.Tag      = $file.FullName

      $pb = New-Object System.Windows.Forms.PictureBox
      $pb.Width     = $thumbSize
      $pb.Height    = $thumbSize
      $pb.Left      = $thumbPad
      $pb.Top       = $thumbPad
      $pb.SizeMode  = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
      $pb.BackColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
      $pb.Cursor    = [System.Windows.Forms.Cursors]::Hand
      $pb.Tag       = $file.FullName   # path only; image loaded lazily

      $lbl = New-Object System.Windows.Forms.Label
      $lbl.Text      = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
      $lbl.Width     = $thumbSize + $thumbPad * 2 - 4
      $lbl.Height    = 20
      $lbl.Left      = 2
      $lbl.Top       = $thumbPad + $thumbSize + 2
      $lbl.ForeColor = [System.Drawing.Color]::Silver
      $lbl.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
      $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
      $lbl.Cursor    = [System.Windows.Forms.Cursors]::Hand
      $lbl.Tag       = $file.FullName

      $onEnter = { param($s,$e) $s.Parent.BackColor = [System.Drawing.Color]::FromArgb(70,70,90) }
      $onLeave = { param($s,$e) $s.Parent.BackColor = [System.Drawing.Color]::FromArgb(45,45,45) }
      $pb.Add_MouseEnter($onEnter);   $pb.Add_MouseLeave($onLeave)
      $lbl.Add_MouseEnter($onEnter);  $lbl.Add_MouseLeave($onLeave)
      $tile.Add_MouseEnter($onEnter); $tile.Add_MouseLeave($onLeave)

      $applyAndClose = {
        param($s,$e)
        $path = $s.Tag
        if (-not $path -and $s.Parent) { $path = $s.Parent.Tag }
        try {
          Set-WallpaperStyleRegistry -Style $script:cfg.Style
          Apply-Wallpaper -ImagePath $path
          Write-Host "Setting wallpaper: $path (Style: $($script:cfg.Style))"
        } catch {
          [System.Windows.Forms.MessageBox]::Show(
            "Failed to set wallpaper:`n$($_.Exception.Message)",
            "Random Wallpaper",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
        }
        $pickerForm.Close()
      }
      $pb.Add_Click($applyAndClose)
      $lbl.Add_Click($applyAndClose)
      $tile.Add_Click($applyAndClose)

      $tile.Controls.Add($pb)
      $tile.Controls.Add($lbl)
      return $tile
    }

    # Build all tiles (no images yet)
    $allTiles = [System.Collections.Generic.List[System.Windows.Forms.Panel]]::new()
    foreach ($f in ($imageFiles | Sort-Object Name)) {
      $tile = & $buildTile $f
      $allTiles.Add($tile)
      $flow.Controls.Add($tile)
    }

    # Lazy-load images in batches via a timer so the form appears immediately
    $lazyState  = @{ Index = 0 }
    $batchSize  = 15
    $loadTimer  = New-Object System.Windows.Forms.Timer
    $loadTimer.Interval = 30

    # Pre-load PresentationCore once if any webp files are present
    if ($imageFiles | Where-Object { $_.Extension.ToLowerInvariant() -eq '.webp' } | Select-Object -First 1) {
      try { Add-Type -AssemblyName PresentationCore -ErrorAction SilentlyContinue } catch {}
    }

    $loadTimer.add_Tick({
      $end = [Math]::Min($lazyState.Index + $batchSize, $allTiles.Count)
      for ($i = $lazyState.Index; $i -lt $end; $i++) {
        $pb = $allTiles[$i].Controls | Where-Object { $_ -is [System.Windows.Forms.PictureBox] } | Select-Object -First 1
        if ($pb -and $pb.Tag -and -not $pb.Image) {
          $imgPath = [string]$pb.Tag
          $ext = [System.IO.Path]::GetExtension($imgPath).ToLowerInvariant()
          if ($ext -eq '.webp') {
            try {
              $stream = [System.IO.File]::OpenRead($imgPath)
              try {
                $decoder = [System.Windows.Media.Imaging.BitmapDecoder]::Create(
                  $stream,
                  [System.Windows.Media.Imaging.BitmapCreateOptions]::None,
                  [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                )
                $frame  = $decoder.Frames[0]
                $enc    = New-Object System.Windows.Media.Imaging.BmpBitmapEncoder
                $enc.Frames.Add($frame)
                $ms = New-Object System.IO.MemoryStream
                $enc.Save($ms)
                $ms.Position = 0
                $pb.Image = [System.Drawing.Bitmap]::new($ms)
                $ms.Dispose()
              } finally { $stream.Dispose() }
            } catch { try { $pb.LoadAsync($imgPath) } catch {} }
          } else {
            try { $pb.LoadAsync($imgPath) } catch {}
          }
        }
      }
      $lazyState.Index = $end
      if ($lazyState.Index -ge $allTiles.Count) { $loadTimer.Stop() }
    })

    $searchBox.Add_TextChanged({
      $q = $searchBox.Text.Trim()
      $flow.SuspendLayout()
      $flow.Controls.Clear()
      foreach ($tile in $allTiles) {
        if ($q -eq '' -or $tile.Tag -like "*$q*") { $flow.Controls.Add($tile) }
      }
      $flow.ResumeLayout()
    })

    $pickerForm.Add_Shown({
      $searchBox.Focus()
      $loadTimer.Start()
    })

    $pickerForm.Add_FormClosed({
      $loadTimer.Stop()
      $loadTimer.Dispose()
      & $disposeAll
      foreach ($tile in $allTiles) { try { $tile.Dispose() } catch {} }
    })

    $pickerForm.ShowDialog() | Out-Null
    $pickerForm.Dispose()
  })

  $itemSetFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $script:cfg.ImageFolder
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $script:cfg.ImageFolder = $dlg.SelectedPath
      Save-Config -Config $script:cfg
    }
  })

  $itemAddFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description  = "Select an extra image folder to add"
    $dlg.SelectedPath = $script:cfg.ImageFolder
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $selected = $dlg.SelectedPath
      if (-not $script:cfg.ExtraFolders) { $script:cfg.ExtraFolders = @() }
      $existing = @($script:cfg.ExtraFolders)
      if ($existing -notcontains $selected -and $selected -ne $script:cfg.ImageFolder) {
        $script:cfg.ExtraFolders = $existing + $selected
        Save-Config -Config $script:cfg
        [System.Windows.Forms.MessageBox]::Show("Added extra folder:`n$selected", "Random Wallpaper") | Out-Null
      }
    }
  })

  $itemClearExtra.Add_Click({
    $script:cfg.ExtraFolders = @()
    Save-Config -Config $script:cfg
    [System.Windows.Forms.MessageBox]::Show("Extra folders cleared.", "Random Wallpaper") | Out-Null
  })

  $itemSetDayFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Folder to use during daytime (6am-8pm)"
    $dlg.SelectedPath = if ($script:cfg.DayFolder -and (Test-Path -LiteralPath $script:cfg.DayFolder)) { $script:cfg.DayFolder } else { $script:cfg.ImageFolder }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $script:cfg.DayFolder = $dlg.SelectedPath
      Save-Config -Config $script:cfg
    }
  })

  $itemSetNightFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Folder to use at night (8pm-6am)"
    $dlg.SelectedPath = if ($script:cfg.NightFolder -and (Test-Path -LiteralPath $script:cfg.NightFolder)) { $script:cfg.NightFolder } else { $script:cfg.ImageFolder }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
      $script:cfg.NightFolder = $dlg.SelectedPath
      Save-Config -Config $script:cfg
    }
  })

  $itemClearTOD.Add_Click({
    $script:cfg.DayFolder   = $null
    $script:cfg.NightFolder = $null
    Save-Config -Config $script:cfg
    [System.Windows.Forms.MessageBox]::Show("Time-of-day folders cleared.", "Random Wallpaper") | Out-Null
  })

  $itemSetInterval.Add_Click({
    $current = [string]$script:cfg.IntervalMinutes
    try {
      $inputVal = [Microsoft.VisualBasic.Interaction]::InputBox("Change every N minutes:", "Random Wallpaper", $current)
    } catch {
      $form = New-Object System.Windows.Forms.Form
      $form.Text = "Random Wallpaper"; $form.Width = 320; $form.Height = 140
      $tb = New-Object System.Windows.Forms.TextBox
      $tb.Text = $current; $tb.Left = 12; $tb.Top = 12; $tb.Width = 280
      $ok     = New-Object System.Windows.Forms.Button
      $ok.Text = "OK";     $ok.Left = 140; $ok.Top = 50; $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
      $cancel = New-Object System.Windows.Forms.Button
      $cancel.Text = "Cancel"; $cancel.Left = 220; $cancel.Top = 50; $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
      $form.Controls.AddRange(@($tb, $ok, $cancel))
      $form.AcceptButton = $ok; $form.CancelButton = $cancel
      $inputVal = if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $tb.Text } else { "" }
      $form.Dispose()
    }
    if ($inputVal -and $inputVal.Trim().Length -gt 0) {
      $parsed = 0
      if ([int]::TryParse($inputVal, [ref]$parsed)) {
        if ($parsed -lt 1) { $parsed = 1 }   # enforce minimum
        $script:cfg.IntervalMinutes = $parsed
        Save-Config -Config $script:cfg
        Update-Timer
      } else {
        [System.Windows.Forms.MessageBox]::Show("Please enter a valid integer number of minutes.", "Random Wallpaper") | Out-Null
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
      [System.Windows.Forms.MessageBox]::Show("Failed to update startup shortcut: $($_.Exception.Message)", "Random Wallpaper") | Out-Null
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
    if ($script:generatedIconHandle -and $script:generatedIconHandle -ne [System.IntPtr]::Zero) {
      try { [Win32.IconHelper]::DestroyIcon($script:generatedIconHandle) | Out-Null } catch {}
    }
    $context.ExitThread()
  })

  # Apply one immediately on start
  try {
    $folders = Get-EffectiveFolders -Config $script:cfg
    Set-RandomWallpaperOnce -Folders $folders -IncludeSubfolders:$script:cfg.Recurse -Style $script:cfg.Style
  } catch {}

  [System.Windows.Forms.Application]::Run($context)
}

function Set-RandomWallpaperOnce {
  param([string[]]$Folders, [switch]$IncludeSubfolders, [string]$Style)
  $image = Get-RandomImagePath -Folders $Folders -IncludeSubfolders:$IncludeSubfolders
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
  $argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File',("`"{0}`"" -f $ScriptPath))
  if ($ResolvedImageFolder) { $argList += @('-ImageFolder',("`"{0}`"" -f $ResolvedImageFolder)) }
  if ($IntervalMinutes)     { $argList += @('-IntervalMinutes',("{0}" -f $IntervalMinutes)) }
  if ($Style)               { $argList += @('-Style',("{0}" -f $Style)) }
  if ($Recurse)             { $argList += '-Recurse' }

  $action    = New-ScheduledTaskAction -Execute 'pwsh.exe' -Argument ($argList -join ' ')
  $trigger1  = New-ScheduledTaskTrigger -AtLogOn
  $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Limited
  $task      = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger1 `
               -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable)
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

# ── Entry point ────────────────────────────────────────────────────────────────
$scriptPath = $PSCommandPath
if (-not $scriptPath) { $scriptPath = Join-Path $PSScriptRoot 'rand_wallpaper.ps1' }

if ($UninstallBackground) { Uninstall-BackgroundTask -TaskName $TaskName; return }
if ($InstallBackground) {
  $resolvedFolderForTask = Resolve-DefaultImageFolder -InputFolder $ImageFolder
  Install-BackgroundTask -ScriptPath $scriptPath -ResolvedImageFolder $resolvedFolderForTask `
    -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse -TaskName $TaskName
  return
}
if ($InstallStartup)   { Install-StartupShortcut -ScriptPath $scriptPath; return }
if ($UninstallStartup) { Uninstall-StartupShortcut; return }

if ($Tray) {
  Start-TrayApp -InitialImageFolder $ImageFolder -IntervalMinutes $IntervalMinutes -Style $Style -Recurse:$Recurse
  return
}

$resolvedFolder = Resolve-DefaultImageFolder -InputFolder $ImageFolder
$useRecurse = if ($PSBoundParameters.ContainsKey('Recurse')) { [bool]$Recurse } else { $true }
Set-RandomWallpaperOnce -Folders @($resolvedFolder) -IncludeSubfolders:$useRecurse -Style $Style

if ($Once) { return }

if ($IntervalMinutes -lt 1) { $IntervalMinutes = 1 }
while ($true) {
  try {
    Start-Sleep -Seconds ($IntervalMinutes * 60)
    Set-RandomWallpaperOnce -Folders @($resolvedFolder) -IncludeSubfolders:$useRecurse -Style $Style
  } catch {
    Write-Warning ("Failed to set wallpaper: " + $_.Exception.Message)
  }
}
