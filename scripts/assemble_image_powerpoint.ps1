[CmdletBinding()]
param([Parameter(Mandatory)][string]$PlanPath,[Parameter(Mandatory)][string]$ManifestPath,[string]$OutputPath)
. "$PSScriptRoot/common.ps1"
. "$PSScriptRoot/resolve_output_path.ps1"
$plan=Read-Json $PlanPath; $manifest=Read-Json $ManifestPath
$root=Split-Path -Parent (Resolve-Path -LiteralPath $ManifestPath).Path
$spec=$plan.presentation_spec
foreach ($key in @('speaker','presentation_date')) { if ([string]::IsNullOrWhiteSpace((Value-Or $spec $key ''))) {throw "Missing cover metadata: $key"} }
$output=Resolve-DeckOutput $OutputPath ((Value-Or $plan 'deck_title' 'Image-Presentation')+'-图片版')
if ([IO.Path]::GetExtension($output) -ne '.pptx' -or (Test-Path -LiteralPath $output)) {throw 'Use a new .pptx output'}
$registry=@{}
foreach ($a in $manifest.assets) {
 if ($a.asset_kind -ne 'slide') {continue}
 if ($a.origin -ne 'imagegen' -or [string]::IsNullOrWhiteSpace($a.prompt) -or [string]::IsNullOrWhiteSpace($a.generation_record)) {throw 'ImageGen full-page provenance required'}
 if ($registry.ContainsKey($a.id) -or [IO.Path]::IsPathRooted($a.path)) {throw 'Invalid asset ID/path'}
 $path=[IO.Path]::GetFullPath((Join-Path $root $a.path))
 if (-not $path.StartsWith($root+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {throw 'Asset outside project'}
 $item=Get-Item -LiteralPath $path
 while ($null -ne $item -and $item.FullName.Length -ge $root.Length) {
  if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {throw 'Reparse-point asset not allowed'}
  $item=if ($item -is [IO.DirectoryInfo]) {$item.Parent} else {$item.Directory}
 }
 if ([IO.Path]::GetExtension($path).ToLowerInvariant() -notin @('.png','.jpg','.jpeg') -or (Get-FileHash -LiteralPath $path).Hash -ne $a.sha256) {throw 'Asset hash/type mismatch'}
 $registry[$a.id]=@{path=$path;hash=$a.sha256}
}
if ($plan.canvas.width -le 0 -or $plan.canvas.height -le 0 -or @($plan.slides).Count -eq 0) {throw 'Invalid canvas/slides'}
$seen=@{}
foreach ($s in $plan.slides) {
 if ($s.kind -notin @('cover','toc','body','ending') -or -not $registry.ContainsKey($s.asset_id)) {throw 'Invalid slide kind/asset'}
 if ((Value-Or $s 'text_verified' $false) -ne $true -or [string]::IsNullOrWhiteSpace($s.text_transcript)) {throw 'Full-page text must be manually verified'}
 if ($seen.ContainsKey($registry[$s.asset_id].hash)) {throw 'Full-page image reused'}
 $seen[$registry[$s.asset_id].hash]=$true
 if ($s.kind -eq 'cover') {foreach ($key in @('speaker','presentation_date')) {if (-not $s.text_transcript.Contains([string]$spec.$key)) {throw "Cover image transcript missing $key"}}}
 if ($s.kind -eq 'body') {
  if (-not $s.title.StartsWith([string]$s.chapter_id+'、') -or -not $s.text_transcript.Contains([string]$s.subtitle) -or -not $s.text_transcript.Contains([string]$s.title)) {throw 'Chapter/title not in verified image text'}
 }
}
[void][IO.Directory]::CreateDirectory((Split-Path -Parent $output))
$app=$null;$deck=$null
try {
 $app=New-Object -ComObject PowerPoint.Application;$deck=$app.Presentations.Add(0)
 $deck.PageSetup.SlideWidth=[single]$plan.canvas.width;$deck.PageSetup.SlideHeight=[single]$plan.canvas.height
 $i=0
 foreach ($s in $plan.slides) {
  $i++;$slide=$deck.Slides.Add($i,12);$slide.DisplayMasterShapes=0
  $pic=$slide.Shapes.AddPicture($registry[$s.asset_id].path,0,-1,0,0,-1,-1)
  $origW=$pic.Width;$origH=$pic.Height
  $ratio=$origW/$origH; $target=$plan.canvas.width/$plan.canvas.height
  if ([Math]::Abs($ratio/$target-1) -gt 0.01) {throw "Page $i image aspect differs; regenerate before assembly"}
  $factor=[Math]::Max($plan.canvas.width/$origW,$plan.canvas.height/$origH)
  $crop=$pic.PictureFormat.Crop;$crop.ShapeLeft=0;$crop.ShapeTop=0
  $crop.ShapeWidth=[single]$plan.canvas.width;$crop.ShapeHeight=[single]$plan.canvas.height
  $crop.PictureWidth=[single]($origW*$factor);$crop.PictureHeight=[single]($origH*$factor)
  $crop.PictureOffsetX=0;$crop.PictureOffsetY=0
  $pic.Name="full-page-$i";$pic.AlternativeText=$s.text_transcript
  $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text="Image-only slide. Core: $($s.core_message)`r`n$($s.notes)"
 }
 $deck.SaveAs($output,24)
 @{output=$output;slides=$i;output_mode='image';sha256=(Get-FileHash -LiteralPath $output).Hash} | ConvertTo-Json
} finally {if ($null -ne $deck) {$deck.Close();Release-Com $deck};Release-Com $app}
