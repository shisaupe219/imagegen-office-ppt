[CmdletBinding()]
param()
. "$PSScriptRoot/common.ps1"
if ($env:OS -ne 'Windows_NT') { throw 'Windows and Microsoft PowerPoint required' }
if ($null -eq [type]::GetTypeFromProgID('PowerPoint.Application')) { throw 'PowerPoint COM not registered' }
$app = $null; $deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Add(0)
    [pscustomobject]@{ powerpoint_version = $app.Version; can_create_presentation = $true; native_charts_supported_by_script = $false } | ConvertTo-Json
} finally {
    if ($null -ne $deck) { $deck.Close(); Release-Com $deck }
    Release-Com $app
}
