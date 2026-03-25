$t = $null; $e = $null
[void][System.Management.Automation.Language.Parser]::ParseFile(
    'c:\Users\magnus.sandstrom\Code\Repo\Random_wallpaper\rand_wallpaper.ps1',
    [ref]$t, [ref]$e)
if ($e) { $e } else { 'No syntax errors' }
