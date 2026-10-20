# Windows PowerShell 5.1. No administrator privileges or execution-policy bypass.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$lock = $null
$stage = $null
try {
    $release = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'release.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($release.schema -ne 1 -or $release.sha256 -notmatch '^[a-f0-9]{64}$') { throw 'Invalid release manifest.' }
    if ($release.package -cne ('package-' + $release.sha256 + '.zip')) { throw 'Invalid package name.' }
    if ($release.launcher -notmatch '^[^\\/:*?"<>|]+\.exe$') { throw 'Invalid launcher name.' }
    if (-not $env:LOCALAPPDATA) { throw 'LOCALAPPDATA is unavailable.' }
    $base = [IO.Path]::GetFullPath($env:LOCALAPPDATA)
    if ($base.StartsWith('\\') -or ([IO.DriveInfo]::new([IO.Path]::GetPathRoot($base))).DriveType -ne 'Fixed') {
        throw 'Local installation requires a fixed local disk.'
    }
    $root = Join-Path $base 'QuoteAuto'
    $versions = Join-Path $root 'versions'
    New-Item -ItemType Directory -Force -Path $versions | Out-Null
    $lock = [IO.File]::Open((Join-Path $root 'install.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
    $target = Join-Path $versions $release.sha256
    $exe = Join-Path $target $release.launcher
    if (-not (Test-Path -LiteralPath (Join-Path $target '.installed'))) {
        if (Test-Path -LiteralPath $target) { throw 'Incomplete version exists; keep current installation and contact support.' }
        $stage = Join-Path $versions ('.install-' + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $stage | Out-Null
        $zip = Join-Path $stage 'payload.zip'
        Write-Host 'Copying this release to your PC (only required when installing/updating)...'
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $release.package) -Destination $zip
        if ((Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant() -cne $release.sha256) {
            throw 'Package checksum mismatch. Existing installation has not been changed.'
        }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $content = Join-Path $stage 'app'
        $prefix = [IO.Path]::GetFullPath($content) + [IO.Path]::DirectorySeparatorChar
        $archive = [IO.Compression.ZipFile]::OpenRead($zip)
        try {
            foreach ($entry in $archive.Entries) {
                $name = $entry.FullName.Replace('/', '\')
                if ($name.Contains(':') -or [IO.Path]::IsPathRooted($name)) { throw 'Unsafe archive path.' }
                $dest = [IO.Path]::GetFullPath((Join-Path $content $name))
                if (-not $dest.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe archive path.' }
            }
        } finally { $archive.Dispose() }
        [IO.Compression.ZipFile]::ExtractToDirectory($zip, $content)
        if (-not (Test-Path -LiteralPath (Join-Path $content $release.launcher) -PathType Leaf)) {
            throw 'Launcher missing from package.'
        }
        Set-Content -LiteralPath (Join-Path $content '.installed') -Value $release.sha256 -Encoding ASCII
        Move-Item -LiteralPath $content -Destination $target
    }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw 'Installed launcher is missing.' }
    if ((Get-Content -LiteralPath (Join-Path $target '.installed') -Raw).Trim() -cne $release.sha256) {
        throw 'Installed version marker is invalid.'
    }
    $shell = New-Object -ComObject WScript.Shell
    $menu = Join-Path ([Environment]::GetFolderPath('Programs')) 'QuoteAuto'
    New-Item -ItemType Directory -Force -Path $menu | Out-Null
    foreach ($folder in @([Environment]::GetFolderPath('Desktop'), $menu)) {
        $shortcut = $shell.CreateShortcut((Join-Path $folder 'QuoteAuto.lnk'))
        $shortcut.TargetPath = $exe
        $shortcut.WorkingDirectory = $target
        $shortcut.Description = 'QuoteAuto - local application'
        $shortcut.Save()
    }
    Write-Host 'Installed. Start QuoteAuto using the desktop or Start menu shortcut.'
    Write-Host 'Cloud documents stay in their original locations. Run Install.cmd again only to update.'
} catch {
    Write-Error -Message $_ -ErrorAction Continue
    exit 1
} finally {
    if ($stage -and (Test-Path -LiteralPath $stage)) { Remove-Item -LiteralPath $stage -Recurse -Force }
    if ($lock) { $lock.Dispose() }
}
