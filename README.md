# Random Wallpaper Scripts

This folder contains two PowerShell scripts for setting a random desktop wallpaper on Windows and configuring it to run automatically at logon.

- rand_wallpaper.ps1 — Sets a random wallpaper immediately and optionally continues changing it on an interval. Can also attempt to set the lock screen image to match.
- Create-RandomWallpaperStartup.ps1 — Creates a Startup entry so rand_wallpaper.ps1 launches automatically at user logon with your chosen options.

## Requirements
- Windows 10/11
- PowerShell (pwsh or Windows PowerShell)
- Images located in a folder you control (jpg, jpeg, png, bmp)

## rand_wallpaper.ps1

Features:
- Immediately sets a random wallpaper when run
- Optional continuous mode with a configurable interval
- Supports styles: Fill, Fit, Stretch, Center, Tile, Span
- Can recurse into subfolders
- Optional: set lock screen image to the same picture (best effort without admin)
- Optional: install/uninstall a background Scheduled Task (may be blocked by policy)

Default image folder resolution order:
1) %USERPROFILE%\wallpapers
2) %USERPROFILE%\Pictures\Wallpapers
3) %USERPROFILE%\Pictures

Usage examples:
- Set once and exit using defaults
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -Once

- Set once and set lock screen too
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -Once -LockScreen

- Run continuously every 15 minutes, include subfolders, style Fill
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -IntervalMinutes 15 -Style Fill -Recurse

- Use a custom folder
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -ImageFolder "C:\Wallpapers" -IntervalMinutes 20 -Style Fit

Parameters:
- -ImageFolder <path>  Folder to pick images from (default logic above)
- -IntervalMinutes <int>  Minutes between changes when running continuously (default 30)
- -Style <Fill|Fit|Stretch|Center|Tile|Span>  Wallpaper style (default Fill)
- -Recurse  Include images from subfolders
- -Once  Set a wallpaper once and exit (no loop)
- -LockScreen  Attempt to set the lock screen image to match
- -InstallBackground  Install a per-user background job (Scheduled Task) that runs at logon
- -UninstallBackground  Remove the background task
- -TaskName <name>  Name for the Scheduled Task (default RandomWallpaper)

Notes on lock screen (Windows 10/11):
- The script writes the lock screen image to C:\Windows\Web\Screen\RandomWallpaper.jpg (preferred) and falls back to C:\ProgramData\RandomWallpaper\RandomWallpaper.jpg if the Windows folder is not writable.
- When run as Administrator, it sets both classic policy and PersonalizationCSP device-wide keys so Windows enforces the image:
  - HKLM\SOFTWARE\Policies\Microsoft\Windows\Personalization\LockScreenImage = C:\Windows\Web\Screen\RandomWallpaper.jpg
  - HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP\LockScreenImagePath = C:\Windows\Web\Screen\RandomWallpaper.jpg
  - HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP\LockScreenImageUrl = C:\Windows\Web\Screen\RandomWallpaper.jpg
  - HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP\LockScreenImageStatus = 1
- The image file is granted read access to ALL APPLICATION PACKAGES, SYSTEM, and Users to ensure LogonUI can read it.
- If not running as Administrator, the script applies best-effort user-level settings; success may vary depending on edition/policies.

Notes on background Scheduled Task:
- The script tries to create a per-user task that launches at logon and loops internally.
- On some managed systems, registering tasks may be blocked; in that case, use the Startup shortcut approach below.

## Create-RandomWallpaperStartup.ps1

Creates a Startup entry so rand_wallpaper.ps1 starts automatically at user logon with your chosen options. It tries to create a .lnk shortcut; if that fails, it creates a .cmd launcher as a fallback.

Typical usage:
- Create a Startup entry to run every 15 minutes, Fill, include subfolders, and update lock screen:
  pwsh -NoProfile -File .\Create-RandomWallpaperStartup.ps1 -IntervalMinutes 15 -Style Fill -Recurse -LockScreen
  Note: For reliable lock screen updates, run rand_wallpaper.ps1 once as Administrator with -LockScreen to set HKLM policy/CSP. After that, the Startup entry can maintain changes.

Parameters:
- -ScriptPath <path>  Path to rand_wallpaper.ps1 (defaults to sibling file in the same folder)
- -IntervalMinutes <int>  Interval in minutes between changes (default 15)
- -Style <Fill|Fit|Stretch|Center|Tile|Span>  Wallpaper style (default Fill)
- -Recurse  Include subfolders
- -LockScreen  Also update lock screen
- -TaskName <name>  Base name for Startup items (default RandomWallpaper)

What it creates:
- Preferred: .lnk shortcut at %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\RandomWallpaper.lnk
- Fallback: .cmd launcher at the same location if .lnk creation is unavailable

To remove autostart:
- Delete RandomWallpaper.lnk and/or RandomWallpaper.cmd from
  %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

## Tips
- Ensure your images exist in %USERPROFILE%\wallpapers (or supply -ImageFolder). Supported: jpg, jpeg, png, bmp.
- For multi-monitor spanning, use -Style Span.
- If lock screen doesn’t change without admin, run an elevated terminal once and use -LockScreen to set the Group Policy value, or rely on desktop wallpaper only.

## Troubleshooting
- "No valid image folder found": Create %USERPROFILE%\wallpapers or pass -ImageFolder to an existing folder with images.
- "Access is denied" when installing background task: Use Create-RandomWallpaperStartup.ps1 instead; some environments restrict Scheduled Task creation.
- Lock screen unchanged:
  - Ensure policy/CSP values exist and point to C:\Windows\Web\Screen\RandomWallpaper.jpg (see Notes on lock screen).
  - Run once as Administrator with -LockScreen to set HKLM keys and ACLs.
  - Force policy update with: gpupdate /target:computer /force
  - Ensure Windows Spotlight is disabled for lock screen in Settings.

