[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PresentationPath,
    [Parameter(Mandatory)][string]$OutputDirectory,
    [ValidateRange(640,3840)][int]$Width = 1920
)
. "$PSScriptRoot/common.ps1"
$inputFile = (Resolve-Path -LiteralPath $PresentationPath).Path
$destination = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $destination) { throw 'Export directory already exists; use a new version' }
[void][IO.Directory]::CreateDirectory($destination)
$app = $null; $deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Open($inputFile, -1, 0, 0)
    $height = [int][Math]::Round($Width * $deck.PageSetup.SlideHeight / $deck.PageSetup.SlideWidth)
    foreach ($slide in $deck.Slides) {
        $slide.Export((Join-Path $destination ('slide-{0:D3}.png' -f $slide.SlideIndex)), 'PNG', $Width, $height)
    }
    @{source_sha256=(Get-FileHash -LiteralPath $inputFile).Hash; slide_count=$deck.Slides.Count; width=$Width; height=$height} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $destination 'export.json') -Encoding UTF8
    [pscustomobject]@{directory=$destination; slides=$deck.Slides.Count; width=$Width; height=$height} | ConvertTo-Json
} finally {
    if ($null -ne $deck) { $deck.Close(); Release-Com $deck }
    Release-Com $app
}
