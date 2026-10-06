[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ManifestPath,
    [Parameter(Mandatory)][string]$AssetPath,
    [Parameter(Mandatory)][string]$Id,
    [string]$Prompt,
    [Parameter(Mandatory)][string]$Purpose,
    [string]$GenerationRecord,
    [ValidateSet('imagegen','user-provided')][string]$Origin = 'imagegen',
    [string]$Source
)
. "$PSScriptRoot/common.ps1"
if ($Id -notmatch '^[a-zA-Z0-9_-]+$') { throw 'Asset ID must contain letters, numbers, underscore or hyphen' }
$manifestFull = [IO.Path]::GetFullPath($ManifestPath)
$root = Split-Path -Parent $manifestFull
$file = (Resolve-Path -LiteralPath $AssetPath).Path
if (-not $file.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Copy asset inside project first' }
$data = if (Test-Path -LiteralPath $manifestFull) { Read-Json $manifestFull } else { [pscustomobject]@{assets=@()} }
if (@($data.assets | Where-Object id -eq $Id).Count) { throw 'Asset ID already registered; use a new version ID' }
if ($Origin -eq 'imagegen' -and ([string]::IsNullOrWhiteSpace($Prompt) -or [string]::IsNullOrWhiteSpace($GenerationRecord))) { throw 'ImageGen prompt and generation record required' }
if ($Origin -eq 'user-provided' -and [string]::IsNullOrWhiteSpace($Source)) { throw 'User asset source required' }
$entry = [pscustomobject]@{id=$Id; path=$file.Substring($root.Length+1).Replace('\','/'); sha256=(Get-FileHash -LiteralPath $file).Hash; origin=$Origin; prompt=$Prompt; purpose=$Purpose; generation_record=$GenerationRecord; source=$Source}
$data.assets = @($data.assets) + $entry
$temp = "$manifestFull.$([Guid]::NewGuid().ToString('N')).tmp"
try {
    $data | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $temp -Encoding UTF8
    Move-Item -LiteralPath $temp -Destination $manifestFull -Force
} finally {
    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp }
}
"Registered $Id; provenance must be reviewed against generation record"
