# Random Wallpaper Scripts

This folder contains two PowerShell scripts for setting a random desktop wallpaper on Windows and configuring it to run automatically at logon.

- rand_wallpaper.ps1 — Sets a random wallpaper immediately and optionally continues changing it on an interval.
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
- Recurses into subfolders by default
- Optional: install/uninstall a background Scheduled Task (may be blocked by policy)

Default image folder resolution order:
1) %USERPROFILE%\wallpapers (default)
2) %USERPROFILE%\Pictures\Wallpapers
3) %USERPROFILE%\Pictures

Usage examples:
- Set once and exit using defaults
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -Once

- Run continuously every 15 minutes, include subfolders, style Fill
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -IntervalMinutes 15 -Style Fill -Recurse

- Use a custom folder
  pwsh -NoProfile -File .\rand_wallpaper.ps1 -ImageFolder "C:\Wallpapers" -IntervalMinutes 20 -Style Fit

Parameters:
- -ImageFolder cpathe  Folder to pick images from (defaults to %USERPROFILE%\wallpapers; see resolution order above)
- -IntervalMinutes <int>  Minutes between changes when running continuously (default 30)
- -Style <Fill|Fit|Stretch|Center|Tile|Span>  Wallpaper style (default Fill)
- -Recurse  Include images from subfolders (default behavior)
- -Once  Set a wallpaper once and exit (no loop)
- -InstallBackground  Install a per-user background job (Scheduled Task) that runs at logon
- -UninstallBackground  Remove the background task
- -TaskName <name>  Name for the Scheduled Task (default RandomWallpaper)

Notes on background Scheduled Task:
- The script tries to create a per-user task that launches at logon and loops internally.
- On some managed systems, registering tasks may be blocked; in that case, use the Startup shortcut approach below.

## Create-RandomWallpaperStartup.ps1

Creates a Startup entry so rand_wallpaper.ps1 starts automatically at user logon with your chosen options. It tries to create a .lnk shortcut; if that fails, it creates a .cmd launcher as a fallback.

PowerShell host selection:
- Prefers Windows PowerShell (powershell.exe) for better hidden/minimized window behavior; falls back to pwsh.exe if not available.
- The startup script relies on rand_wallpaper.ps1's default ImageFolder (%USERPROFILE%\wallpapers). If you want a different folder, run rand_wallpaper.ps1 manually with -ImageFolder to test, or modify the default path in the script.

Typical usage:
- Create a Startup entry to run every 15 minutes, Fill, include subfolders:
  pwsh -NoProfile -File .\Create-RandomWallpaperStartup.ps1 -IntervalMinutes 15 -Style Fill -Recurse

Parameters:
- -ScriptPath \u003cpath\u003e  Path to rand_wallpaper.ps1 (defaults to sibling file in the same folder)
- -IntervalMinutes \u003cint\u003e  Interval in minutes between changes (default 15)
- -Style \u003cFill|Fit|Stretch|Center|Tile|Span\u003e  Wallpaper style (default Fill)
- -Recurse  Include subfolders
- -TaskName \u003cname\u003e  Base name for Startup items (default RandomWallpaper)

What it creates:
- Preferred: .lnk shortcut at %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\RandomWallpaper.lnk
- Fallback: .cmd launcher at the same location if .lnk creation is unavailable. The .cmd uses the same PowerShell host and arguments as the shortcut.

To remove autostart:
- Delete RandomWallpaper.lnk and/or RandomWallpaper.cmd from
  %APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

## Tips
- Ensure your images exist in %USERPROFILE%\wallpapers (or supply -ImageFolder). Supported: jpg, jpeg, png, bmp.
- For multi-monitor spanning, use -Style Span.

## Troubleshooting
- "No valid image folder found": Create %USERPROFILE%\wallpapers or pass -ImageFolder to an existing folder with images.
- "Access is denied" when installing background task: Use Create-RandomWallpaperStartup.ps1 instead; some environments restrict Scheduled Task creation.
