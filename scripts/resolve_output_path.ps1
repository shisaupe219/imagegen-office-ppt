function Resolve-DeckOutput([string]$OutputPath,[string]$Name='presentation') {
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { return [IO.Path]::GetFullPath($OutputPath) }
    $desktop = [Environment]::GetFolderPath('Desktop')
    if ([string]::IsNullOrWhiteSpace($desktop) -or -not (Test-Path -LiteralPath $desktop -PathType Container)) { throw 'System Desktop unavailable; specify OutputPath' }
    foreach ($c in [IO.Path]::GetInvalidFileNameChars()) { $Name=$Name.Replace([string]$c,'_') }
    $Name=$Name.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($Name)) { $Name='presentation' }
    if ($Name.Length -gt 90) { $Name=$Name.Substring(0,90) }
    $candidate=Join-Path $desktop "$Name.pptx"; $version=2
    while (Test-Path -LiteralPath $candidate) { $candidate=Join-Path $desktop "$Name-v$version.pptx"; $version++ }
    return $candidate
}
