# Random Wallpaper (Windows)

A PowerShell script that sets a random desktop wallpaper from a folder. It supports a Windows system tray icon with a right‑click menu to manually change the wallpaper, configure the schedule (minutes), choose the image folder, toggle subfolder inclusion, pick style, and enable/disable scheduling. Settings persist under your profile.

## Features
- System tray app (Windows Forms)
  - Change now
  - Set folder… (browse)
  - Set interval (minutes)…
  - Include subfolders (toggle)
  - Style: Fill, Fit, Stretch, Center, Tile, Span
  - Scheduling enabled (toggle)
  - Start at logon (toggle via Startup shortcut)
  - Open images folder
  - Exit
- Custom tray icon
  - Uses `tray.ico` (or `icon.ico`) if present next to the script
  - If none is present, the script auto-generates a multi-resolution `tray.ico` (16, 24, 32, 48, 64, 128, 256) with an image‑themed design (mountain + sun)
  - The tray and Startup shortcut will use the appropriate size for your DPI
- One‑off or continuous (non‑tray) modes
- Optional Scheduled Task to run continuously at user logon
- Supported images: `jpg`, `jpeg`, `png`, `bmp`

## Requirements
- Windows 10/11
- PowerShell 7+ (tested with 7.5.3), executable `pwsh`
- .NET Windows Desktop assemblies (System.Windows.Forms/System.Drawing) – present by default on Windows desktops

## Files
- `rand_wallpaper.ps1` – the main script
- `tray.ico` – auto-generated on first run if missing; place your own to override (or use `icon.ico`)
- Config: `%AppData%\RandomWallpaper\config.json` (created/updated by tray mode)

## Quick start
Run the tray app (recommended):
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -Tray
```
Then right‑click the tray icon to configure your folder, schedule, recurse, and style. Double‑click the tray icon or select “Change now” to immediately set a random wallpaper.

## Autostart at sign‑in (Startup shortcut)
Install a Startup shortcut that launches the tray app at logon:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -InstallStartup
```
Remove it:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -UninstallStartup
```
You can also toggle “Start at logon” directly from the tray menu.

Notes:
- The shortcut is created in your user Startup folder with target `pwsh.exe` and arguments `-NoProfile -ExecutionPolicy Bypass -File "<script>" -Tray`.
- The shortcut icon uses `tray.ico` or `icon.ico` if available; otherwise a system icon.

## Scheduled Task (optional, non‑tray)
Create a Scheduled Task to start the script at logon and run continuously (the script itself loops on your selected interval):
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -InstallBackground -ImageFolder "C:\Wallpapers" -IntervalMinutes 30 -Style Fill -Recurse -TaskName RandomWallpaper
```
Remove the task:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -UninstallBackground -TaskName RandomWallpaper
```
Notes:
- Uninstall may prompt for elevation.
- The task is registered for the current user (Interactive, Limited).

## One‑off change
Set one random wallpaper and exit:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -Once
```

## Continuous (non‑tray) session
Run in the foreground continuously, changing every N minutes:
```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -ImageFolder "C:\Wallpapers" -IntervalMinutes 30 -Style Fill -Recurse
```
Press Ctrl+C to stop.

## Parameters (CLI)
- `-ImageFolder <path>`
  - Folder containing images. Default prefers `%USERPROFILE%\Pictures\Wallpapers` (or Known Folders variant), otherwise falls back to `%USERPROFILE%\Pictures` and `%USERPROFILE%\wallpapers`.
- `-IntervalMinutes <int>`
  - Minutes between changes. Default: `240`. Minimum enforced: `1`.
- `-Style <Fill|Fit|Stretch|Center|Tile|Span>`
  - Wallpaper style. Default: `Stretch`.
- `-Recurse` (switch)
  - Include images from subfolders. Default: `true` unless explicitly set otherwise.
- `-Once` (switch)
  - Set a random wallpaper once and exit.
- `-Tray` (switch)
  - Start the tray app (Windows Forms message loop). The tray has its own scheduler/timer and persists settings.
- `-InstallStartup` / `-UninstallStartup` (switch)
  - Install/remove a Startup shortcut that launches tray mode at sign‑in.
- `-InstallBackground` / `-UninstallBackground` (switch)
  - Install/remove a Scheduled Task that starts the script at logon and runs continuously.
- `-TaskName <string>`
  - Name for the Scheduled Task. Default: `RandomWallpaper`.

## Tray mode details
- Persisted settings: `%AppData%\RandomWallpaper\config.json`
  - `ImageFolder`: string
  - `IntervalMinutes`: number (minutes)
  - `Style`: one of `Fill|Fit|Stretch|Center|Tile|Span`
  - `Recurse`: boolean
  - `ScheduleEnabled`: boolean (enables the tray timer)
- Timer/scheduler is internal to the tray app and can be toggled via the menu (“Scheduling enabled”).
- Double‑click the tray icon = “Change now”.
- Icon:
  - Provide `tray.ico` (or `icon.ico`) next to the script for a custom icon. A 16×16 (and/or multi‑resolution) ICO is recommended.
  - If not present, a small image‑style icon is drawn programmatically.
- If you don’t see the tray icon, check the hidden icons overflow menu in the Windows taskbar.

## Styles mapping
The script sets registry values equivalent to Windows styles:
- Fill → WallpaperStyle=10, TileWallpaper=0
- Fit → WallpaperStyle=6, TileWallpaper=0
- Stretch → WallpaperStyle=2, TileWallpaper=0
- Center → WallpaperStyle=0, TileWallpaper=0
- Tile → WallpaperStyle=0, TileWallpaper=1
- Span → WallpaperStyle=22, TileWallpaper=0 (multi‑monitor spanning)

## Supported image formats
- `jpg`, `jpeg`, `png`, `bmp`

## Troubleshooting
- “No images found”: Check the selected folder and whether “Include subfolders” is set appropriately.
- Execution policy blocks script: use the examples here which pass `-ExecutionPolicy Bypass` per‑invocation, or `Unblock-File .\\rand_wallpaper.ps1` as needed.
- Tray icon not visible: It may be hidden in the taskbar overflow. Drag it into the visible area if desired.
- Multi‑monitor behavior: Use `Span` style to span a single image across displays; other styles apply per‑monitor.
- OneDrive Pictures redirection: The script resolves the Windows Known Folder for Pictures to find your default `Wallpapers` folder if present.
- Error “SystemParametersInfo failed…”: Ensure the image path is accessible and the file type is supported.

## Uninstall / Cleanup
- Remove Startup shortcut:
  ```powershell
  pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -UninstallStartup
  ```
- Remove Scheduled Task:
  ```powershell
  pwsh -NoProfile -ExecutionPolicy Bypass -File .\rand_wallpaper.ps1 -UninstallBackground -TaskName RandomWallpaper
  ```
- Remove tray settings: delete `%AppData%\RandomWallpaper\config.json` (optional). Close the tray app via its “Exit” menu.

## License
Not specified.
