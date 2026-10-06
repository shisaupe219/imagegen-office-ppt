[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PreviewDirectory,
    [Parameter(Mandatory)][string]$OutputDirectory,
    [ValidateRange(1,100)][int]$Limit = 20,
    [ValidateRange(1,6)][int]$Columns = 4
)
. "$PSScriptRoot/common.ps1"
$source = (Resolve-Path -LiteralPath $PreviewDirectory).Path
if (-not (Test-Path -LiteralPath (Join-Path $source 'export.json'))) { throw 'Use a PowerPoint export directory containing export.json' }
$images = @(Get-ChildItem -LiteralPath $source -Filter 'slide-*.png' | Sort-Object Name | Select-Object -First $Limit)
if (-not $images.Count) { throw 'No exported slides found' }
$dest = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $dest) { throw 'Output directory exists; use a new version' }
[void][IO.Directory]::CreateDirectory($dest)
$app = $null; $deck = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Add(0)
    $rows = [int][Math]::Ceiling($images.Count / $Columns)
    $cellWidth = 912 / $Columns
    $pictureWidth = $cellWidth-16
    $cellHeight = $pictureWidth*9/16 + 24
    $sheetHeight = 48+$rows*$cellHeight
    if ($sheetHeight -gt 4032) { throw 'Overview exceeds PowerPoint height limit; use more columns or split pages' }
    $deck.PageSetup.SlideWidth = 960; $deck.PageSetup.SlideHeight = [single]$sheetHeight
    $slide = $deck.Slides.Add(1,12)
    $slide.DisplayMasterShapes = 0; $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid(); $slide.Background.Fill.ForeColor.RGB = Color-Value 'FFFFFF'
    for ($i=0; $i -lt $images.Count; $i++) {
        $x = 24 + ($i % $Columns)*$cellWidth; $y = 24 + [Math]::Floor($i/$Columns)*$cellHeight
        $pic = $slide.Shapes.AddPicture($images[$i].FullName,0,-1,0,0,-1,-1)
        $pic.LockAspectRatio = -1
        $factor = [Math]::Min($pictureWidth/$pic.Width,($cellHeight-24)/$pic.Height)
        $pic.Width = [single]($pic.Width*$factor)
        $pic.Left = [single]($x+($pictureWidth-$pic.Width)/2); $pic.Top = [single]$y
    }
    $slide.Export((Join-Path $dest 'contact-sheet.png'),'PNG',2560,[int][Math]::Round(2560*$sheetHeight/960))
    @{source_export=(Read-Json (Join-Path $source 'export.json')); slides=$images.Name} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $dest 'contact-sheet.json') -Encoding UTF8
    'PowerPoint contact sheet exported'
} finally {
    if ($null -ne $deck) { $deck.Close(); Release-Com $deck }
    Release-Com $app
}
