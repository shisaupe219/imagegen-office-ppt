[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PlanPath,
    [Parameter(Mandatory)][string]$ManifestPath,
    [Parameter(Mandatory)][string]$OutputPath
)
. "$PSScriptRoot/common.ps1"
$plan = Read-Json $PlanPath
$manifestFull = (Resolve-Path -LiteralPath $ManifestPath).Path
$root = Split-Path -Parent $manifestFull
$manifest = Read-Json $manifestFull
$output = [IO.Path]::GetFullPath($OutputPath)
if ([IO.Path]::GetExtension($output) -ne '.pptx') { throw 'Output must be .pptx' }
if (Test-Path -LiteralPath $output) { throw 'Output already exists; use a versioned filename' }
if (Test-Path -LiteralPath "$output.audit.json") { throw 'Audit already exists; use a versioned filename' }
if ($plan.canvas.width -le 0 -or $plan.canvas.height -le 0) { throw 'Invalid canvas' }
Check-Box $plan.title_style $plan.canvas
Check-Box $plan.page_style $plan.canvas
if ($plan.page_style.x -lt ($plan.canvas.width * 0.75) -or $plan.page_style.y -lt ($plan.canvas.height * 0.85)) { throw 'Page number must be bottom-right' }
$assets = @{}
foreach ($asset in $manifest.assets) {
    if ($assets.ContainsKey($asset.id)) { throw "Duplicate asset ID: $($asset.id)" }
    if ($asset.id -notmatch '^[a-zA-Z0-9_-]+$') { throw 'Invalid asset ID' }
    if ($asset.origin -ne 'imagegen' -or [string]::IsNullOrWhiteSpace($asset.prompt) -or [string]::IsNullOrWhiteSpace((Value-Or $asset 'generation_record' ''))) { throw 'Asset must have ImageGen provenance, prompt and generation record' }
    if ([IO.Path]::IsPathRooted($asset.path)) { throw 'Asset paths must be relative' }
    $path = [IO.Path]::GetFullPath((Join-Path $root $asset.path))
    if (-not $path.StartsWith($root + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Asset outside project root' }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing asset: $path" }
    # Reject reparse points in the entire path, including intermediate directories.
    $part = Get-Item -LiteralPath $path
    if ($part.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Reparse-point assets not supported' }
    $dir = (Get-Item -LiteralPath $path).Directory
    while ($null -ne $dir -and $dir.FullName.Length -ge $root.Length) {
        if ($dir.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Reparse-point directories not supported' }
        $dir = $dir.Parent
    }
    if ([IO.Path]::GetExtension($path).ToLowerInvariant() -notin @('.png','.jpg','.jpeg')) { throw 'PNG/JPEG assets required' }
    if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ne $asset.sha256) { throw "Asset hash mismatch: $($asset.id)" }
    $assets[$asset.id] = $path
}
if (@($plan.slides).Count -eq 0) { throw 'No slides' }
foreach ($s in $plan.slides) {
    if ($s.kind -notin @('cover','toc','body','ending')) { throw 'Unknown slide kind' }
    if ([string]::IsNullOrWhiteSpace($s.title) -or [string]::IsNullOrWhiteSpace($s.core_message)) { throw 'Title and core_message required' }
    foreach ($e in $s.elements) {
        Check-Box $e $plan.canvas
        if ($e.type -notin @('text','image')) { throw "Unsupported element: $($e.type)" }
        if ($e.type -eq 'image') {
            if (-not $assets.ContainsKey($e.asset_id)) { throw "Unregistered asset: $($e.asset_id)" }
            if ((Value-Or $e 'fit' 'contain') -notin @('contain','cover')) { throw 'Image fit must be contain or cover' }
        }
    }
}
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $output))
$app = $null; $deck = $null; $audit = @(); $overflow = @()
try {
    $app = New-Object -ComObject PowerPoint.Application
    $deck = $app.Presentations.Add(0)
    $deck.PageSetup.SlideWidth = [single]$plan.canvas.width
    $deck.PageSetup.SlideHeight = [single]$plan.canvas.height
    $defaultFont = [string](Value-Or $plan.title_style 'font' 'Microsoft YaHei')
    $page = 0
    foreach ($s in $plan.slides) {
        $page++
        $slide = $deck.Slides.Add($page, 12)
        $slide.FollowMasterBackground = 0
        $slide.DisplayMasterShapes = 0
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = Color-Value (Value-Or $plan.canvas 'background' 'FFFFFF')
        $elementIndex = 0
        foreach ($e in $s.elements) {
            $elementIndex++
            if ($e.type -eq 'text') {
                $shape = Add-Text $slide $e.text $e $defaultFont "text-$elementIndex"
            } else {
                $shape = $slide.Shapes.AddPicture($assets[$e.asset_id], 0, -1, 0, 0, -1, -1)
                $shape.Name = "asset-$($e.asset_id)-$elementIndex"
                $origW = [double]$shape.Width; $origH = [double]$shape.Height
                $shape.LockAspectRatio = -1
                $factor = if ((Value-Or $e 'fit' 'contain') -eq 'cover') { [Math]::Max($e.w/$origW, $e.h/$origH) } else { [Math]::Min($e.w/$origW, $e.h/$origH) }
                if ((Value-Or $e 'fit' 'contain') -eq 'cover') {
                    $crop = $shape.PictureFormat.Crop
                    $crop.ShapeLeft = [single]$e.x; $crop.ShapeTop = [single]$e.y
                    $crop.ShapeWidth = [single]$e.w; $crop.ShapeHeight = [single]$e.h
                    $crop.PictureWidth = [single]($origW*$factor); $crop.PictureHeight = [single]($origH*$factor)
                    $crop.PictureOffsetX = 0; $crop.PictureOffsetY = 0
                } else {
                    $shape.Width = [single]($origW*$factor)
                    $shape.Left = [single]($e.x + ($e.w-$shape.Width)/2)
                    $shape.Top = [single]($e.y + ($e.h-$shape.Height)/2)
                }
                $shape.AlternativeText = "ImageGen asset: $($e.asset_id)"
            }
        }
        $titleBox = if ($s.kind -eq 'body') { $plan.title_style } else { Value-Or $s 'title_box' $plan.title_style }
        Check-Box $titleBox $plan.canvas
        [void](Add-Text $slide $s.title $titleBox $defaultFont 'slide-title')
        [void](Add-Text $slide ([string]$page) $plan.page_style $defaultFont 'page-number')
        $notes = "Core message: $($s.core_message)`r`nEvidence: $(@($s.evidence) -join '; ')`r`nSources: $(@($s.sources) -join '; ')"
        $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes
        $items = @()
        foreach ($obj in $slide.Shapes) {
            $items += @{name=$obj.Name; left=$obj.Left; top=$obj.Top; width=$obj.Width; height=$obj.Height}
            if ($obj.HasTextFrame -eq -1 -and $obj.TextFrame.HasText -eq -1) {
                if ($obj.TextFrame.TextRange.BoundHeight -gt ($obj.Height + 1) -or $obj.TextFrame.TextRange.BoundWidth -gt ($obj.Width + 1)) { $overflow += "slide-$page/$($obj.Name)" }
            }
        }
        $audit += @{page=$page; kind=$s.kind; shapes=$items}
    }
    $deck.SaveAs($output, 24)
    @{presentation_sha256=(Get-FileHash -LiteralPath $output).Hash; slides=$audit; possible_text_overflow=$overflow; source_manifest=$manifestFull} | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath "$output.audit.json" -Encoding UTF8
    [pscustomobject]@{output=$output; slides=$page; possible_text_overflow=$overflow} | ConvertTo-Json -Depth 4
} finally {
    if ($null -ne $deck) { $deck.Close(); Release-Com $deck }
    Release-Com $app
}
